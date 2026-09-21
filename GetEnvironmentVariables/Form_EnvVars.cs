using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Sunny.UI;
using System.Collections;
using System.Diagnostics;
using System.Security.Principal;



namespace GetEnvironmentVariables
{
    #region DirectoryContextClass
    public class DirectoryContext
    {
        #region DirectoryProperties
        public string ExcelOutputs { get; }
        public string JsonOutputs { get; }
        public string PowerShellOutputs { get; }
        #endregion



        #region DirectoryContextConstructor
        public DirectoryContext()
        {
            // Identify Contexts
            string environment = Environment.GetEnvironmentVariable("DOTNET_ENVIRONMENT") ?? "Production";
            string machineName = Environment.MachineName; // "VJ_MUSICDEV" on your home PC

            // Build the Configuration Cascade
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddJsonFile($"appsettings.{machineName}.json", optional: true) // Overrides if Home PC
                .AddJsonFile($"appsettings.{environment}.json", optional: true) // Overrides if inside Visual Studio
                .AddJsonFile("appsettings.Local.json", optional: true)          // Overrides if local dev file exists
                .Build();

            // Resolve and Assign Paths
            ExcelOutputs = ResolvePath(config["Directories:Excel"]);
            JsonOutputs = ResolvePath(config["Directories:JSON"]);
            PowerShellOutputs = ResolvePath(config["Directories:PowerShell"]);
        }
        #endregion



        #region HelperMethod
        private string ResolvePath(string? configuredPath)
        {
            if (string.IsNullOrWhiteSpace(configuredPath))
                throw new InvalidOperationException("A required directory path is missing from the appsettings configuration.");

            // If it's an absolute path (G:\..., B:\..., C:\...), return it as-is.
            // If it's a relative path ("ExcelOutputs"), combine it with the exe's current directory.
            return Path.IsPathRooted(configuredPath)
                ? configuredPath
                : Path.Combine(AppContext.BaseDirectory, configuredPath);
        }
        #endregion
    }
    #endregion



    #region FormClass
    public class Form_EnvVars : UIForm
    {
        #region ClassConfig
        const int FORM_WIDTH = 1080;
        const int FORM_HEIGHT = 800;
        const int GRID_MARGIN = 16;
        const int HEADER_HEIGHT = 36;
        const int TOAST_DURATION_MS = 1200;
        const int NAME_COLUMN_WIDTH = 260;

        private readonly DirectoryContext _directories;
        private readonly Color _bgDark = Color.FromArgb(30, 30, 46);
        private readonly Color _panelBg = Color.FromArgb(36, 36, 54);
        private readonly Color _userAccent = Color.FromArgb(0, 188, 212);
        private readonly Color _systemAccent = Color.FromArgb(255, 179, 0);
        private readonly Color _gridBg = Color.FromArgb(40, 42, 60);
        private readonly Color _gridAltRow = Color.FromArgb(48, 50, 72);
        private readonly Color _gridSelectionBg = Color.FromArgb(60, 64, 98);
        private readonly Color _textPrimary = Color.FromArgb(230, 230, 240);
        private readonly Color _textSecondary = Color.FromArgb(170, 170, 190);
        private readonly Color _headerRowBg = Color.FromArgb(28, 28, 42);
        private readonly Color _borderSubtle = Color.FromArgb(58, 60, 82);

        private UIPanel _userPanel;
        private UIPanel _systemPanel;
        private DataGridView _userGrid;
        private DataGridView _systemGrid;
        private UISymbolButton _btnExport;
        private UISymbolButton _btnAddUser;
        private UISymbolButton _btnAddSystem;
        private Label _toastLabel;
        private System.Windows.Forms.Timer _toastTimer;
        private ContextMenuStrip _gridContextMenu;
        private DataGridView _contextTargetGrid;
        private int _contextTargetRow = -1;

        private List<EnvVarEntry> _userVars = new();
        private List<EnvVarEntry> _systemVars = new();
        #endregion



        #region FormConstructor
        public Form_EnvVars()
        {
            _directories = new DirectoryContext();

            InitializeForm();
            InitializeToast();
            InitializeContextMenu();
            InitializeUserPanel();
            InitializeSystemPanel();
            InitializeExportButton();
            InitializeEvents();
        }

        private void InitializeForm()
        {
            // Form properties
            this.Text = "Environment Variables";
            this.Size = new Size(FORM_WIDTH, FORM_HEIGHT);
            this.MinimumSize = new Size(860, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = _bgDark;
            this.AllowShowTitle = true;
            this.ShowTitle = true;

            // Sunny.UI style
            this.Style = UIStyle.DarkBlue;
            this.RectColor = _borderSubtle;
            this.ForeColor = _textPrimary;
            this.TitleColor = _userAccent;
            this.TitleFont = new Font("Segoe UI Semibold", 11f);
        }

        private void InitializeToast()
        {
            // Floating toast label for copy confirmation
            _toastLabel = new Label
            {
                AutoSize = false,
                Size = new Size(140, 32),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI Semibold", 10f),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(220, 46, 204, 113),
                Visible = false,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };

            _toastTimer = new System.Windows.Forms.Timer { Interval = TOAST_DURATION_MS };
            _toastTimer.Tick += (s, e) =>
            {
                _toastTimer.Stop();
                _toastLabel.Visible = false;
            };

            this.Controls.Add(_toastLabel);
            _toastLabel.BringToFront();
        }

        private void InitializeContextMenu()
        {
            // Styled right-click context menu for Update
            _gridContextMenu = new ContextMenuStrip();
            _gridContextMenu.BackColor = Color.FromArgb(40, 42, 60);
            _gridContextMenu.ForeColor = _textPrimary;
            _gridContextMenu.Font = new Font("Segoe UI", 10f);
            _gridContextMenu.ShowImageMargin = false;
            _gridContextMenu.Renderer = new DarkMenuRenderer(_borderSubtle, _gridSelectionBg);

            ToolStripMenuItem updateItem = new ToolStripMenuItem("✏️  Update Variable");
            updateItem.Click += ContextMenu_UpdateClick;
            _gridContextMenu.Items.Add(updateItem);
        }

        private void InitializeEvents()
        {
            this.Load += Form_EnvVars_Load;
            this.Resize += Form_EnvVars_Resize;
        }
        #endregion



        #region FormLoad
        private void Form_EnvVars_Load(object? sender, EventArgs e)
        {
            // Epplus License Context
            ExcelPackage.License.SetNonCommercialPersonal("<Vince Jackson>");

            // Retrieve environment variables
            _userVars = GetScopedEnvVars(EnvironmentVariableTarget.User);
            _systemVars = GetScopedEnvVars(EnvironmentVariableTarget.Machine);

            // Populate grids
            PopulateGrid(_userGrid, _userVars, "User");
            PopulateGrid(_systemGrid, _systemVars, "System");

            // Position everything
            LayoutPanels();
        }

        private void Form_EnvVars_Resize(object? sender, EventArgs e)
        {
            LayoutPanels();
        }
        #endregion



        #region EnvironmentVariableRetrieval
        private List<EnvVarEntry> GetScopedEnvVars(EnvironmentVariableTarget target)
        {
            List<EnvVarEntry> results = new();

            try
            {
                IDictionary rawVars = Environment.GetEnvironmentVariables(target);

                foreach (DictionaryEntry entry in rawVars)
                {
                    string name = entry.Key?.ToString() ?? string.Empty;
                    string value = entry.Value?.ToString() ?? string.Empty;

                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        results.Add(new EnvVarEntry(name, value));
                    }
                }

                results.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to retrieve {target} environment variables:\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            return results;
        }
        #endregion



        #region PanelConstruction
        private void InitializeUserPanel()
        {
            _userPanel = CreateSectionPanel("User Environment Variables", _userAccent);
            _userGrid = CreateStyledGrid(_userAccent);
            _userGrid.CellClick += Grid_CellClick;
            _userGrid.CellMouseDown += Grid_CellMouseDown;
            _userPanel.Controls.Add(_userGrid);

            // Add button inside User panel header
            _btnAddUser = CreateAddButton(_userAccent);
            _btnAddUser.Click += (s, e) => HandleAddVariable(EnvironmentVariableTarget.User);
            _userPanel.Controls.Add(_btnAddUser);
            _btnAddUser.BringToFront();

            this.Controls.Add(_userPanel);
        }

        private void InitializeSystemPanel()
        {
            _systemPanel = CreateSectionPanel("System Environment Variables", _systemAccent);
            _systemGrid = CreateStyledGrid(_systemAccent);
            _systemGrid.CellClick += Grid_CellClick;
            _systemGrid.CellMouseDown += Grid_CellMouseDown;
            _systemPanel.Controls.Add(_systemGrid);

            // Add button inside System panel header
            _btnAddSystem = CreateAddButton(_systemAccent);
            _btnAddSystem.Click += (s, e) => HandleAddVariable(EnvironmentVariableTarget.Machine);
            _systemPanel.Controls.Add(_btnAddSystem);
            _btnAddSystem.BringToFront();

            this.Controls.Add(_systemPanel);
        }

        private UIPanel CreateSectionPanel(string title, Color accentColor)
        {
            UIPanel panel = new UIPanel
            {
                Style = UIStyle.Custom,
                FillColor = _panelBg,
                RectColor = _borderSubtle,
                Radius = 8,
                RadiusSides = UICornerRadiusSides.All,
                Font = new Font("Segoe UI Semibold", 11f),
                ForeColor = accentColor,
                Text = $"  {title}"
            };

            return panel;
        }

        private UISymbolButton CreateAddButton(Color accentColor)
        {
            UISymbolButton btn = new UISymbolButton
            {
                Text = " Add",
                Symbol = 61543, // Plus icon
                SymbolSize = 16,
                Size = new Size(82, 28),
                Style = UIStyle.Custom,
                FillColor = Color.FromArgb(44, 48, 70),
                FillHoverColor = Color.FromArgb(56, 62, 92),
                FillPressColor = Color.FromArgb(38, 40, 60),
                RectColor = accentColor,
                ForeColor = _textPrimary,
                SymbolColor = accentColor,
                Font = new Font("Segoe UI Semibold", 9f),
                Radius = 5,
                Cursor = Cursors.Hand
            };

            return btn;
        }

        private DataGridView CreateStyledGrid(Color accentColor)
        {
            DataGridView grid = new DataGridView
            {
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToOrderColumns = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                EnableHeadersVisualStyles = false,
                BackgroundColor = _gridBg,
                GridColor = _borderSubtle,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _gridBg,
                    ForeColor = _textPrimary,
                    SelectionBackColor = _gridSelectionBg,
                    SelectionForeColor = Color.White,
                    Font = new Font("Cascadia Code", 9.5f, FontStyle.Regular),
                    Padding = new Padding(6, 4, 6, 4)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _gridAltRow,
                    ForeColor = _textPrimary,
                    SelectionBackColor = _gridSelectionBg,
                    SelectionForeColor = Color.White,
                    Font = new Font("Cascadia Code", 9.5f, FontStyle.Regular),
                    Padding = new Padding(6, 4, 6, 4)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _headerRowBg,
                    ForeColor = accentColor,
                    Font = new Font("Segoe UI Semibold", 10f),
                    Alignment = DataGridViewContentAlignment.MiddleLeft,
                    Padding = new Padding(8, 6, 6, 6),
                    SelectionBackColor = _headerRowBg,
                    SelectionForeColor = accentColor
                },
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 38,
                RowTemplate = { Height = 30 },
                ScrollBars = ScrollBars.Vertical,
                Cursor = Cursors.Hand
            };

            // Add columns
            DataGridViewTextBoxColumn colName = new DataGridViewTextBoxColumn
            {
                Name = "colName",
                HeaderText = "Name",
                Width = NAME_COLUMN_WIDTH,
                MinimumWidth = 160,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };

            DataGridViewTextBoxColumn colValue = new DataGridViewTextBoxColumn
            {
                Name = "colValue",
                HeaderText = "Value",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                MinimumWidth = 200,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };

            grid.Columns.AddRange(colName, colValue);

            // Tooltip for UX
            grid.ShowCellToolTips = true;
            grid.CellFormatting += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                {
                    grid.Rows[e.RowIndex].Cells[e.ColumnIndex].ToolTipText = "Click to copy  |  Right-click to edit";
                }
            };

            return grid;
        }
        #endregion



        #region GridPopulation
        private void PopulateGrid(DataGridView grid, List<EnvVarEntry> vars, string scope)
        {
            grid.Rows.Clear();

            if (vars.Count == 0)
            {
                // Edge case: no environment variables found
                grid.Rows.Add("(none)", $"No {scope} environment variables detected.");
                grid.Rows[0].DefaultCellStyle.ForeColor = _textSecondary;
                grid.Rows[0].DefaultCellStyle.Font = new Font("Segoe UI", 9.5f, FontStyle.Italic);
                return;
            }

            foreach (EnvVarEntry entry in vars)
            {
                grid.Rows.Add(entry.Name, entry.Value);
            }
        }

        private void RefreshGridData(EnvironmentVariableTarget target)
        {
            if (target == EnvironmentVariableTarget.User)
            {
                _userVars = GetScopedEnvVars(EnvironmentVariableTarget.User);
                PopulateGrid(_userGrid, _userVars, "User");
            }
            else
            {
                _systemVars = GetScopedEnvVars(EnvironmentVariableTarget.Machine);
                PopulateGrid(_systemGrid, _systemVars, "System");
            }
        }
        #endregion



        #region LayoutAndPositioning
        private void LayoutPanels()
        {
            if (_userPanel == null || _systemPanel == null) return;

            int clientTop = this.TitleHeight > 0 ? this.TitleHeight : 35;
            int availableWidth = this.ClientSize.Width - (GRID_MARGIN * 2);
            int buttonBarHeight = 56;
            int availableHeight = this.ClientSize.Height - clientTop - (GRID_MARGIN * 3) - buttonBarHeight;
            int panelHeight = (availableHeight - GRID_MARGIN) / 2;

            // User panel — top
            _userPanel.Location = new Point(GRID_MARGIN, clientTop + GRID_MARGIN);
            _userPanel.Size = new Size(availableWidth, panelHeight);

            // System panel — bottom
            _systemPanel.Location = new Point(GRID_MARGIN, clientTop + GRID_MARGIN + panelHeight + GRID_MARGIN);
            _systemPanel.Size = new Size(availableWidth, panelHeight);

            // Position grids inside panels
            LayoutGridInPanel(_userGrid, _userPanel);
            LayoutGridInPanel(_systemGrid, _systemPanel);

            // Position Add buttons inside panel header areas (top-right)
            LayoutAddButton(_btnAddUser, _userPanel);
            LayoutAddButton(_btnAddSystem, _systemPanel);

            // Export button
            if (_btnExport != null)
            {
                _btnExport.Location = new Point(
                    this.ClientSize.Width - _btnExport.Width - GRID_MARGIN,
                    this.ClientSize.Height - _btnExport.Height - GRID_MARGIN);
            }

            // Toast label
            if (_toastLabel != null)
            {
                _toastLabel.Location = new Point(
                    this.ClientSize.Width - _toastLabel.Width - GRID_MARGIN - 8,
                    this.ClientSize.Height - _toastLabel.Height - GRID_MARGIN - 60);
            }
        }

        private void LayoutGridInPanel(DataGridView grid, UIPanel panel)
        {
            int gridTop = HEADER_HEIGHT + 4;
            int gridPad = 10;
            grid.Location = new Point(gridPad, gridTop);
            grid.Size = new Size(panel.Width - (gridPad * 2), panel.Height - gridTop - gridPad);
        }

        private void LayoutAddButton(UISymbolButton btn, UIPanel panel)
        {
            if (btn == null || panel == null) return;
            btn.Location = new Point(panel.Width - btn.Width - 12, 4);
        }
        #endregion



        #region ClickToCopy
        private void Grid_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (sender is not DataGridView grid) return;

            object? cellValue = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            if (cellValue == null) return;

            string text = cellValue.ToString() ?? string.Empty;
            if (string.IsNullOrEmpty(text) || text == "(none)") return;

            try
            {
                Clipboard.SetText(text);
                ShowToast("✓  Copied!");
            }
            catch
            {
                ShowToast("Copy failed");
            }
        }

        private void ShowToast(string message)
        {
            _toastTimer.Stop();
            _toastLabel.Text = message;
            _toastLabel.Visible = true;
            _toastLabel.BringToFront();
            _toastTimer.Start();
        }
        #endregion



        #region RightClickUpdate
        private void Grid_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            if (e.RowIndex < 0) return;
            if (sender is not DataGridView grid) return;

            // Check for the empty placeholder row
            object? nameVal = grid.Rows[e.RowIndex].Cells["colName"].Value;
            if (nameVal != null && nameVal.ToString() == "(none)") return;

            // Select the right-clicked row and store context
            grid.ClearSelection();
            grid.Rows[e.RowIndex].Selected = true;
            _contextTargetGrid = grid;
            _contextTargetRow = e.RowIndex;

            // Show context menu at cursor
            Rectangle cellRect = grid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
            _gridContextMenu.Show(grid, cellRect.Left + e.X, cellRect.Top + e.Y);
        }

        private void ContextMenu_UpdateClick(object? sender, EventArgs e)
        {
            if (_contextTargetGrid == null || _contextTargetRow < 0) return;

            // Get current values from the right-clicked row
            string currentName = _contextTargetGrid.Rows[_contextTargetRow].Cells["colName"].Value?.ToString() ?? string.Empty;
            string currentValue = _contextTargetGrid.Rows[_contextTargetRow].Cells["colValue"].Value?.ToString() ?? string.Empty;
            EnvironmentVariableTarget target = ResolveTargetFromGrid(_contextTargetGrid);

            // System scope requires admin elevation
            if (target == EnvironmentVariableTarget.Machine && !EnsureAdminElevation())
                return;

            // Open editor in Update mode
            using (Form_EditEnvVar editor = new Form_EditEnvVar(
                EditMode.Update, currentName, currentValue, target))
            {
                if (editor.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        ApplyEnvVarUpdate(currentName, editor.ResultName, editor.ResultValue, target);
                        RefreshGridData(target);
                        ShowToast("✓  Updated!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"Failed to update environment variable:\n{ex.Message}",
                            "Update Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ApplyEnvVarUpdate(string oldName, string newName, string newValue, EnvironmentVariableTarget target)
        {
            bool nameChanged = !string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase);

            if (nameChanged)
            {
                // Rename = delete old + create new
                Environment.SetEnvironmentVariable(oldName, null, target);
            }

            Environment.SetEnvironmentVariable(newName, newValue, target);
        }
        #endregion



        #region AddNewVariable
        private void HandleAddVariable(EnvironmentVariableTarget target)
        {
            // System scope requires admin elevation
            if (target == EnvironmentVariableTarget.Machine && !EnsureAdminElevation())
                return;

            using (Form_EditEnvVar editor = new Form_EditEnvVar(
                EditMode.Add, string.Empty, string.Empty, target))
            {
                if (editor.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        Environment.SetEnvironmentVariable(editor.ResultName, editor.ResultValue, target);
                        RefreshGridData(target);
                        ShowToast("✓  Added!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"Failed to add environment variable:\n{ex.Message}",
                            "Add Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
        }
        #endregion



        #region AdminElevation
        private bool EnsureAdminElevation()
        {
            if (IsRunningAsAdmin())
                return true;

            DialogResult result = MessageBox.Show(
                "Modifying System environment variables requires administrator privileges.\n\n" +
                "The application will relaunch with elevated permissions.\nAny unsaved work will be preserved in Windows.",
                "Admin Privileges Required",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);

            if (result != DialogResult.OK)
            {
                ShowToast("Elevation cancelled");
                return false;
            }

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = Application.ExecutablePath,
                    UseShellExecute = true,
                    Verb = "runas"
                };

                Process.Start(psi);
                Application.Exit();
                return false; // Won't reach here, but compiler needs it
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // User declined UAC prompt
                ShowToast("Elevation cancelled");
                return false;
            }
        }

        private bool IsRunningAsAdmin()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }
        #endregion



        #region ExcelExport
        private void InitializeExportButton()
        {
            _btnExport = new UISymbolButton
            {
                Text = "  Export to Excel",
                Symbol = 61568, // Printer / document icon
                SymbolSize = 20,
                Size = new Size(180, 42),
                Style = UIStyle.Custom,
                FillColor = Color.FromArgb(48, 52, 78),
                FillHoverColor = Color.FromArgb(60, 66, 100),
                FillPressColor = Color.FromArgb(40, 44, 68),
                RectColor = _userAccent,
                ForeColor = _textPrimary,
                SymbolColor = _userAccent,
                Font = new Font("Segoe UI Semibold", 10f),
                Radius = 6,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };

            _btnExport.Click += BtnExport_Click;
            this.Controls.Add(_btnExport);
            _btnExport.BringToFront();
        }

        private void BtnExport_Click(object? sender, EventArgs e)
        {
            try
            {
                // Ensure output directory exists
                if (!Directory.Exists(_directories.ExcelOutputs))
                {
                    Directory.CreateDirectory(_directories.ExcelOutputs);
                }

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string filePath = Path.Combine(_directories.ExcelOutputs, $"EnvVars_{timestamp}.xlsx");

                using (ExcelPackage package = new ExcelPackage())
                {
                    WriteEnvVarSheet(package, "User Variables", _userVars, _userAccent);
                    WriteEnvVarSheet(package, "System Variables", _systemVars, _systemAccent);

                    FileInfo fileInfo = new FileInfo(filePath);
                    package.SaveAs(fileInfo);

                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }

                ShowToast("✓  Exported!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to export Excel file:\n{ex.Message}",
                    "Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void WriteEnvVarSheet(ExcelPackage package, string sheetName, List<EnvVarEntry> vars, Color accentColor)
        {
            ExcelWorksheet ws = package.Workbook.Worksheets.Add(sheetName);

            // Color palette
            Color darkBg = Color.FromArgb(30, 30, 46);
            Color headerBg = Color.FromArgb(28, 28, 42);
            Color rowBg = Color.FromArgb(40, 42, 60);
            Color altRowBg = Color.FromArgb(48, 50, 72);
            Color textWhite = Color.FromArgb(230, 230, 240);
            Color borderColor = Color.FromArgb(58, 60, 82);

            // Title row
            ws.Cells[1, 1].Value = sheetName;
            ws.Cells[1, 1, 1, 2].Merge = true;

            using (ExcelRange titleRange = ws.Cells[1, 1, 1, 2])
            {
                titleRange.Style.Font.Size = 14;
                titleRange.Style.Font.Bold = true;
                titleRange.Style.Font.Color.SetColor(accentColor);
                titleRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                titleRange.Style.Fill.BackgroundColor.SetColor(darkBg);
                titleRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            ws.Row(1).Height = 32;

            // Timestamp row
            ws.Cells[2, 1].Value = "Exported:";
            ws.Cells[2, 2].Value = DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss");
            ws.Cells[2, 1, 2, 2].Style.Font.Size = 9;
            ws.Cells[2, 1, 2, 2].Style.Font.Italic = true;
            ws.Cells[2, 1, 2, 2].Style.Font.Color.SetColor(Color.FromArgb(140, 140, 160));
            ws.Cells[2, 1, 2, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[2, 1, 2, 2].Style.Fill.BackgroundColor.SetColor(darkBg);
            ws.Row(2).Height = 20;

            // Header row
            int headerRow = 3;
            ws.Cells[headerRow, 1].Value = "Name";
            ws.Cells[headerRow, 2].Value = "Value";

            using (ExcelRange hdrRange = ws.Cells[headerRow, 1, headerRow, 2])
            {
                hdrRange.Style.Font.Bold = true;
                hdrRange.Style.Font.Size = 11;
                hdrRange.Style.Font.Color.SetColor(accentColor);
                hdrRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                hdrRange.Style.Fill.BackgroundColor.SetColor(headerBg);
                hdrRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                hdrRange.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                hdrRange.Style.Border.Bottom.Color.SetColor(accentColor);
            }

            ws.Row(headerRow).Height = 28;

            // Data rows
            int dataStartRow = 4;

            if (vars.Count == 0)
            {
                ws.Cells[dataStartRow, 1].Value = "(none)";
                ws.Cells[dataStartRow, 2].Value = "No environment variables detected.";
                ws.Cells[dataStartRow, 1, dataStartRow, 2].Style.Font.Italic = true;
                ws.Cells[dataStartRow, 1, dataStartRow, 2].Style.Font.Color.SetColor(Color.FromArgb(140, 140, 160));
                ws.Cells[dataStartRow, 1, dataStartRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[dataStartRow, 1, dataStartRow, 2].Style.Fill.BackgroundColor.SetColor(rowBg);
            }
            else
            {
                for (int i = 0; i < vars.Count; i++)
                {
                    int row = dataStartRow + i;
                    bool isAlt = i % 2 == 1;
                    Color bg = isAlt ? altRowBg : rowBg;

                    ws.Cells[row, 1].Value = vars[i].Name;
                    ws.Cells[row, 2].Value = vars[i].Value;

                    using (ExcelRange dataRange = ws.Cells[row, 1, row, 2])
                    {
                        dataRange.Style.Font.Size = 10;
                        dataRange.Style.Font.Name = "Cascadia Code";
                        dataRange.Style.Font.Color.SetColor(textWhite);
                        dataRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        dataRange.Style.Fill.BackgroundColor.SetColor(bg);
                        dataRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        dataRange.Style.Border.Bottom.Color.SetColor(borderColor);
                    }

                    ws.Row(row).Height = 22;
                }
            }

            // Column widths
            ws.Column(1).Width = 36;
            ws.Column(2).Width = 90;

            // Freeze panes below header
            ws.View.FreezePanes(dataStartRow, 1);

            // Sheet tab color
            ws.TabColor = accentColor;
        }
        #endregion



        #region HelperMethods
        private EnvironmentVariableTarget ResolveTargetFromGrid(DataGridView grid)
        {
            return grid == _userGrid
                ? EnvironmentVariableTarget.User
                : EnvironmentVariableTarget.Machine;
        }
        #endregion
    }
    #endregion



    #region EditEnvVarDialogClass
    public enum EditMode
    {
        Add,
        Update
    }

    public class Form_EditEnvVar : UIForm
    {
        #region ClassConfig
        const int MAX_NAME_LENGTH = 255;
        const int MAX_VALUE_LENGTH = 32767;

        private static readonly char[] INVALID_NAME_CHARS = { '=', '%', '<', '>', '^', '&', '|', ':' };

        private readonly EditMode _mode;
        private readonly string _originalName;
        private readonly string _originalValue;
        private readonly EnvironmentVariableTarget _target;

        private UITextBox _txtName;
        private UITextBox _txtValue;
        private UIButton _btnAction;
        private UIButton _btnCancel;
        private Label _lblError;

        public string ResultName { get; private set; } = string.Empty;
        public string ResultValue { get; private set; } = string.Empty;
        #endregion



        #region DialogConstructor
        public Form_EditEnvVar(EditMode mode, string currentName, string currentValue, EnvironmentVariableTarget target)
        {
            _mode = mode;
            _originalName = currentName;
            _originalValue = currentValue;
            _target = target;

            InitializeDialog();
            InitializeControls();
        }

        private void InitializeDialog()
        {
            string scopeLabel = _target == EnvironmentVariableTarget.User ? "User" : "System";
            string actionLabel = _mode == EditMode.Update ? "Update" : "Add";

            this.Text = $"{actionLabel} {scopeLabel} Environment Variable";
            this.Size = new Size(480, 290);
            this.MinimumSize = this.Size;
            this.MaximumSize = new Size(600, 290);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;

            // Sunny.UI theming
            this.Style = UIStyle.DarkBlue;
            this.BackColor = Color.FromArgb(30, 30, 46);
            this.RectColor = Color.FromArgb(58, 60, 82);
            this.ForeColor = Color.FromArgb(230, 230, 240);
            this.TitleColor = _target == EnvironmentVariableTarget.User
                ? Color.FromArgb(0, 188, 212)
                : Color.FromArgb(255, 179, 0);
            this.TitleFont = new Font("Segoe UI Semibold", 10f);
            this.AllowShowTitle = true;
            this.ShowTitle = true;
        }

        private void InitializeControls()
        {
            Color accentColor = _target == EnvironmentVariableTarget.User
                ? Color.FromArgb(0, 188, 212)
                : Color.FromArgb(255, 179, 0);

            int titleOffset = this.TitleHeight > 0 ? this.TitleHeight : 35;
            int labelX = 24;
            int fieldX = 24;
            int fieldWidth = this.ClientSize.Width - 48;
            int currentY = titleOffset + 12;

            // Name label
            UILabel lblName = new UILabel
            {
                Text = "Name:",
                Location = new Point(labelX, currentY),
                Size = new Size(fieldWidth, 20),
                Font = new Font("Segoe UI Semibold", 9.5f),
                ForeColor = accentColor,
                Style = UIStyle.Custom
            };
            this.Controls.Add(lblName);
            currentY += 22;

            // Name textbox
            _txtName = new UITextBox
            {
                Text = _originalName,
                Location = new Point(fieldX, currentY),
                Size = new Size(fieldWidth, 32),
                Font = new Font("Cascadia Code", 10f),
                Style = UIStyle.Custom,
                FillColor = Color.FromArgb(40, 42, 60),
                ForeColor = Color.FromArgb(230, 230, 240),
                RectColor = Color.FromArgb(58, 60, 82)
            };
            this.Controls.Add(_txtName);
            currentY += 40;

            // Value label
            UILabel lblValue = new UILabel
            {
                Text = "Value:",
                Location = new Point(labelX, currentY),
                Size = new Size(fieldWidth, 20),
                Font = new Font("Segoe UI Semibold", 9.5f),
                ForeColor = accentColor,
                Style = UIStyle.Custom
            };
            this.Controls.Add(lblValue);
            currentY += 22;

            // Value textbox
            _txtValue = new UITextBox
            {
                Text = _originalValue,
                Location = new Point(fieldX, currentY),
                Size = new Size(fieldWidth, 32),
                Font = new Font("Cascadia Code", 10f),
                Style = UIStyle.Custom,
                FillColor = Color.FromArgb(40, 42, 60),
                ForeColor = Color.FromArgb(230, 230, 240),
                RectColor = Color.FromArgb(58, 60, 82)
            };
            this.Controls.Add(_txtValue);
            currentY += 40;

            // Error label (hidden by default)
            _lblError = new Label
            {
                AutoSize = false,
                Size = new Size(fieldWidth - 170, 20),
                Location = new Point(fieldX, currentY + 6),
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(255, 85, 85),
                BackColor = Color.Transparent,
                Text = string.Empty,
                Visible = false
            };
            this.Controls.Add(_lblError);

            // Action button (Update or Add)
            string actionText = _mode == EditMode.Update ? "Update" : "Add";
            int btnWidth = 90;
            int btnHeight = 34;
            int btnSpacing = 10;
            int btnY = currentY;
            int btnRightEdge = fieldX + fieldWidth;

            _btnCancel = new UIButton
            {
                Text = "Cancel",
                Size = new Size(btnWidth, btnHeight),
                Location = new Point(btnRightEdge - btnWidth, btnY),
                Style = UIStyle.Custom,
                FillColor = Color.FromArgb(48, 50, 72),
                FillHoverColor = Color.FromArgb(60, 64, 90),
                FillPressColor = Color.FromArgb(40, 42, 60),
                RectColor = Color.FromArgb(80, 82, 100),
                ForeColor = Color.FromArgb(200, 200, 210),
                Font = new Font("Segoe UI Semibold", 9.5f),
                Radius = 5,
                Cursor = Cursors.Hand
            };
            _btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
            this.Controls.Add(_btnCancel);

            _btnAction = new UIButton
            {
                Text = actionText,
                Size = new Size(btnWidth, btnHeight),
                Location = new Point(btnRightEdge - (btnWidth * 2) - btnSpacing, btnY),
                Style = UIStyle.Custom,
                FillColor = Color.FromArgb(0, 140, 160),
                FillHoverColor = Color.FromArgb(0, 168, 192),
                FillPressColor = Color.FromArgb(0, 120, 140),
                RectColor = accentColor,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 9.5f),
                Radius = 5,
                Cursor = Cursors.Hand
            };
            _btnAction.Click += BtnAction_Click;
            this.Controls.Add(_btnAction);
        }
        #endregion



        #region Validation
        private void BtnAction_Click(object? sender, EventArgs e)
        {
            string name = _txtName.Text?.Trim() ?? string.Empty;
            string value = _txtValue.Text?.Trim() ?? string.Empty;

            // Validate name is not empty
            if (string.IsNullOrWhiteSpace(name))
            {
                ShowError("Variable name cannot be empty.");
                _txtName.Focus();
                return;
            }

            // Validate value is not empty
            if (string.IsNullOrWhiteSpace(value))
            {
                ShowError("Variable value cannot be empty.");
                _txtValue.Focus();
                return;
            }

            // Validate name does not contain invalid characters
            foreach (char c in INVALID_NAME_CHARS)
            {
                if (name.Contains(c))
                {
                    ShowError($"Name cannot contain '{c}' character.");
                    _txtName.Focus();
                    return;
                }
            }

            // Validate name length
            if (name.Length > MAX_NAME_LENGTH)
            {
                ShowError($"Name cannot exceed {MAX_NAME_LENGTH} characters.");
                _txtName.Focus();
                return;
            }

            // Validate value length
            if (value.Length > MAX_VALUE_LENGTH)
            {
                ShowError($"Value cannot exceed {MAX_VALUE_LENGTH:N0} characters.");
                _txtValue.Focus();
                return;
            }

            // Validate name does not start with a digit
            if (char.IsDigit(name[0]))
            {
                ShowError("Name cannot start with a digit.");
                _txtName.Focus();
                return;
            }

            // Update mode: check for actual changes
            if (_mode == EditMode.Update)
            {
                bool nameChanged = !string.Equals(_originalName, name, StringComparison.OrdinalIgnoreCase);
                bool valueChanged = !string.Equals(_originalValue, value, StringComparison.Ordinal);

                if (!nameChanged && !valueChanged)
                {
                    ShowError("No changes detected.");
                    return;
                }
            }

            // Add mode: check if variable already exists in this scope
            if (_mode == EditMode.Add)
            {
                string? existing = Environment.GetEnvironmentVariable(name, _target);
                if (existing != null)
                {
                    ShowError($"'{name}' already exists in this scope.");
                    _txtName.Focus();
                    return;
                }
            }

            // All validation passed
            ResultName = name;
            ResultValue = value;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void ShowError(string message)
        {
            _lblError.Text = message;
            _lblError.Visible = true;
        }
        #endregion
    }
    #endregion



    #region DarkMenuRendererClass
    public class DarkMenuRenderer : ToolStripProfessionalRenderer
    {
        private readonly Color _borderColor;
        private readonly Color _hoverColor;

        public DarkMenuRenderer(Color borderColor, Color hoverColor)
            : base(new DarkColorTable(borderColor))
        {
            _borderColor = borderColor;
            _hoverColor = hoverColor;
        }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            if (e.Item.Selected)
            {
                using (SolidBrush brush = new SolidBrush(_hoverColor))
                {
                    e.Graphics.FillRectangle(brush, e.Item.ContentRectangle);
                }
            }
            else
            {
                base.OnRenderMenuItemBackground(e);
            }
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            using (Pen pen = new Pen(_borderColor))
            {
                e.Graphics.DrawRectangle(pen, 0, 0, e.AffectedBounds.Width - 1, e.AffectedBounds.Height - 1);
            }
        }
    }

    public class DarkColorTable : ProfessionalColorTable
    {
        private readonly Color _borderColor;

        public DarkColorTable(Color borderColor)
        {
            _borderColor = borderColor;
        }

        public override Color MenuBorder => _borderColor;
        public override Color MenuItemBorder => Color.Transparent;
        public override Color ToolStripDropDownBackground => Color.FromArgb(40, 42, 60);
        public override Color ImageMarginGradientBegin => Color.FromArgb(40, 42, 60);
        public override Color ImageMarginGradientMiddle => Color.FromArgb(40, 42, 60);
        public override Color ImageMarginGradientEnd => Color.FromArgb(40, 42, 60);
        public override Color SeparatorDark => Color.FromArgb(58, 60, 82);
        public override Color SeparatorLight => Color.FromArgb(58, 60, 82);
    }
    #endregion



    #region EnvVarEntryModel
    public class EnvVarEntry
    {
        public string Name { get; }
        public string Value { get; }

        public EnvVarEntry(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }
    #endregion
}