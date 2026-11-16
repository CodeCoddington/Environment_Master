using DanMarDev.DuplicateInstanceCheck;
using DanMarDev.WakeDrives;
using DanMarDev.WinForms_CustomTools;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Sunny.UI;
using System.Collections;
using System.Diagnostics;
using System.Reflection;
using System.Security.Principal;
using static DanMarDev.WinForms_CustomTools.MsgBoxType;



namespace Environment_Master
{
    public class Form_EnvVars : UIForm
    {
        #region ClassConfig
        const int FORM_WIDTH = 1080;
        const int FORM_HEIGHT = 820;
        const int GRID_MARGIN = 16;
        const int HEADER_HEIGHT = 36;
        const int TOAST_DURATION_MS = 1300;
        const int NAME_COLUMN_WIDTH = 260;
        const string CELL_TOOLTIP_TEXT = "Left-click to copy  |  Right-click to update or delete";

        public readonly string MachineName = Environment.MachineName;
        public readonly string TemplateXlsmFilePath = @"D:\My Drive\_GlobalSync_Programming\Excel & Sheets\EnvVars\_Template\EnvVars_Template.xlsm";
        public readonly string XlsmOutputDirectory = @"D:\My Drive\_GlobalSync_Programming\Excel & Sheets\EnvVars\Exports";

        private Color _bgDark { get; } = Color.FromArgb(30, 30, 46);
        private Color _panelBg { get; } = Color.FromArgb(36, 36, 54);
        private Color _userAccent { get; } = Color.FromArgb(0, 188, 212);
        private Color _systemAccent { get; } = Color.FromArgb(255, 179, 0);
        private Color _deleteAccent { get; } = Color.FromArgb(255, 107, 107);
        private Color _gridBg { get; } = Color.FromArgb(40, 42, 60);
        private Color _gridAltRow { get; } = Color.FromArgb(48, 50, 72);
        private Color _gridSelectionBg { get; } = Color.FromArgb(60, 64, 98);
        private Color _textPrimary { get; } = Color.FromArgb(230, 230, 240);
        private Color _textSecondary { get; } = Color.FromArgb(170, 170, 190);
        private Color _headerRowBg { get; } = Color.FromArgb(28, 28, 42);
        private Color _borderSubtle { get; } = Color.FromArgb(58, 60, 82);

        private Font _gridCellFont { get; } = new Font("Cascadia Code", 9.5f, FontStyle.Regular);
        private Font _gridHeaderFont { get; } = new Font("Segoe UI Semibold", 10f);
        private Font _gridEmptyFont { get; } = new Font("Segoe UI", 9.5f, FontStyle.Italic);

        private UIPanel _userPanel { get; set; }
        private UIPanel _systemPanel { get; set; }
        private DataGridView _userGrid { get; set; }
        private DataGridView _systemGrid { get; set; }
        private UITextBox _txtSearch { get; set; }
        private UISymbolButton _btnClearSearch { get; set; }
        private UISymbolButton _btnExport { get; set; }
        private UISymbolButton _btnAddUser { get; set; }
        private UISymbolButton _btnAddSystem { get; set; }
        private Label _toastLabel { get; set; }
        private System.Windows.Forms.Timer _toastTimer { get; set; }
        private ContextMenuStrip _gridContextMenu { get; set; }
        private DataGridView _contextTargetGrid { get; set; }
        private int _contextTargetRow { get; set; } = -1;

        private List<EnvVarEntry> _userVars { get; set; } = new();
        private List<EnvVarEntry> _systemVars { get; set; } = new();
        #endregion



        #region FormConstructor
        public Form_EnvVars()
        {
            InitializeForm();
            InitializeToast();
            InitializeContextMenu();
            InitializeUserPanel();
            InitializeSystemPanel();
            InitializeBottomBar();
            InitializeEvents();
        }

        private void InitializeForm()
        {
            // Configure base form dimensions and dark styling
            this.Text = "Environment Variables";
            this.Size = new Size(FORM_WIDTH, FORM_HEIGHT);
            this.MinimumSize = new Size(860, 620);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = _bgDark;
            this.AllowShowTitle = true;
            this.ShowTitle = true;

            // Apply Sunny.UI theme settings
            this.Style = UIStyle.DarkBlue;
            this.RectColor = _borderSubtle;
            this.ForeColor = _textPrimary;
            this.TitleColor = _userAccent;
            this.TitleFont = new Font("Segoe UI Semibold", 11f);
        }

        private void InitializeToast()
        {
            // Floating toast label for action feedback
            _toastLabel = new Label
            {
                AutoSize = false,
                Size = new Size(150, 32),
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
            // Styled right-click context menu for Update and Delete actions
            _gridContextMenu = new ContextMenuStrip
            {
                BackColor = _gridBg,
                ForeColor = _textPrimary,
                Font = new Font("Segoe UI", 10f),
                ShowImageMargin = false,
                Renderer = new DarkMenuRenderer(_borderSubtle, _gridSelectionBg)
            };

            ToolStripMenuItem updateItem = new ToolStripMenuItem("✏️  Update Variable")
            {
                ForeColor = _textPrimary,
                Padding = new Padding(4, 4, 4, 4)
            };
            updateItem.Click += ContextMenu_UpdateClick;

            ToolStripSeparator separator = new ToolStripSeparator();

            ToolStripMenuItem deleteItem = new ToolStripMenuItem("🗑️  Delete Variable")
            {
                ForeColor = _deleteAccent,
                Padding = new Padding(4, 4, 4, 4)
            };
            deleteItem.Click += ContextMenu_DeleteClick;

            _gridContextMenu.Items.Add(updateItem);
            _gridContextMenu.Items.Add(separator);
            _gridContextMenu.Items.Add(deleteItem);
        }

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
                    Font = _gridCellFont,
                    Padding = new Padding(6, 4, 6, 4)
                },
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _gridAltRow,
                    ForeColor = _textPrimary,
                    SelectionBackColor = _gridSelectionBg,
                    SelectionForeColor = Color.White,
                    Font = _gridCellFont,
                    Padding = new Padding(6, 4, 6, 4)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = _headerRowBg,
                    ForeColor = accentColor,
                    Font = _gridHeaderFont,
                    Alignment = DataGridViewContentAlignment.MiddleLeft,
                    Padding = new Padding(8, 6, 6, 6),
                    SelectionBackColor = _headerRowBg,
                    SelectionForeColor = accentColor
                },
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 38,
                RowTemplate = { Height = 30 },
                ScrollBars = ScrollBars.Vertical,
                ShowCellToolTips = true,
                Cursor = Cursors.Hand
            };

            // Enable double buffering via reflection to prevent scroll and resize flicker
            typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.SetValue(grid, true);

            // Configure Name and Value columns
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

            return grid;
        }

        private UISymbolButton CreateAddButton(Color accentColor)
        {
            UISymbolButton btn = new UISymbolButton
            {
                Text = " Add",
                Symbol = 61543,
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

        private void InitializeBottomBar()
        {
            // Quick filter textbox on bottom-left for searching variables across both scopes
            _txtSearch = new UITextBox
            {
                Watermark = "Filter variables by name or value...",
                Size = new Size(340, 38),
                Font = new Font("Segoe UI", 10f),
                Style = UIStyle.Custom,
                FillColor = _gridBg,
                ForeColor = _textPrimary,
                RectColor = _borderSubtle,
                Radius = 6,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _txtSearch.TextChanged += TxtSearch_TextChanged;
            this.Controls.Add(_txtSearch);

            // Clear filter button next to search box
            _btnClearSearch = new UISymbolButton
            {
                Text = "Clear",
                Symbol = 61453,
                SymbolSize = 15,
                Size = new Size(76, 38),
                Style = UIStyle.Custom,
                FillColor = Color.FromArgb(44, 48, 70),
                FillHoverColor = Color.FromArgb(56, 62, 92),
                FillPressColor = Color.FromArgb(38, 40, 60),
                RectColor = _borderSubtle,
                ForeColor = _textSecondary,
                SymbolColor = _textSecondary,
                Font = new Font("Segoe UI Semibold", 9f),
                Radius = 6,
                Visible = false,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                Cursor = Cursors.Hand
            };
            _btnClearSearch.Click += (s, e) =>
            {
                _txtSearch.Text = string.Empty;
                _txtSearch.Focus();
            };
            this.Controls.Add(_btnClearSearch);

            // Export to Excel button on bottom-right
            _btnExport = new UISymbolButton
            {
                Text = "  Export to Excel",
                Symbol = 61568,
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

        private void InitializeEvents()
        {
            this.Load += Form_EnvVars_Load;
            this.Resize += Form_EnvVars_Resize;
        }
        #endregion



        #region FormLifecycleEvents
        private void Form_EnvVars_Load(object? sender, EventArgs e)
        {
            // Verify single application instance mutex and halt immediately if duplicate
            if (DuplicateInstanceCheck.IsDuplicateAppInstance())
            {
                MsgBox.Show(ErrorExit, "Another instance of this application is already running.");
                Application.Exit();
                return;
            }

            // Wake corporate UNC G: drive when running outside home dev machine
            if (MachineName != "VJ_MUSICDEV")
            {
                if (!WakeDrives.Wake_GDrive().DriveAwake)
                {
                    MsgBox.Show(ErrorExit, "G: Drive is inaccessible.");
                    Application.Exit();
                    return;
                }
            }

            // Configure EPPlus non-commercial license context
            ExcelPackage.License.SetNonCommercialPersonal("<Vince Jackson>");

            // Indicate elevated privileges in form title when running as Administrator
            if (IsRunningAsAdmin())
            {
                this.Text = "Environment Variables  [Administrator]";
            }

            // Retrieve and display User and System environment variables
            _userVars = GetScopedEnvVars(EnvironmentVariableTarget.User);
            _systemVars = GetScopedEnvVars(EnvironmentVariableTarget.Machine);
            ApplySearchFilter();

            // Position panels and controls
            LayoutPanels();
        }

        private void Form_EnvVars_Resize(object? sender, EventArgs e)
        {
            LayoutPanels();
        }
        #endregion



        #region GridAndSearchInteraction
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

        private void Grid_CellMouseDown(object? sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (sender is not DataGridView grid) return;

            // Ignore placeholder row when grid has no items
            object? nameVal = grid.Rows[e.RowIndex].Cells["colName"].Value;
            if (nameVal != null && nameVal.ToString() == "(none)") return;

            // Select right-clicked cell and store target context
            grid.ClearSelection();
            grid.CurrentCell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
            grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = true;
            _contextTargetGrid = grid;
            _contextTargetRow = e.RowIndex;

            // Display context menu at cursor position
            Rectangle cellRect = grid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
            _gridContextMenu.Show(grid, cellRect.Left + e.X, cellRect.Top + e.Y);
        }

        private void TxtSearch_TextChanged(object? sender, EventArgs e)
        {
            _btnClearSearch.Visible = !string.IsNullOrEmpty(_txtSearch.Text);
            ApplySearchFilter();
        }
        #endregion



        #region ContextMenuActions
        private void ContextMenu_UpdateClick(object? sender, EventArgs e)
        {
            if (_contextTargetGrid == null || _contextTargetRow < 0) return;
            if (_contextTargetRow >= _contextTargetGrid.Rows.Count) return;

            // Retrieve variable name, value, and scope from selected row
            string currentName = _contextTargetGrid.Rows[_contextTargetRow].Cells["colName"].Value?.ToString() ?? string.Empty;
            string currentValue = _contextTargetGrid.Rows[_contextTargetRow].Cells["colValue"].Value?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(currentName) || currentName == "(none)") return;

            EnvironmentVariableTarget target = ResolveTargetFromGrid(_contextTargetGrid);

            // Require admin elevation for System scope modifications
            if (target == EnvironmentVariableTarget.Machine && !EnsureAdminElevation())
            {
                return;
            }

            // Launch editor dialog (automatically uses RichTextBox layout for Path/PATH)
            using (Form_EditEnvVar editor = new Form_EditEnvVar(
                EditMode.Update, currentName, currentValue, target))
            {
                if (editor.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        this.UseWaitCursor = true;
                        Cursor.Current = Cursors.WaitCursor;

                        ApplyEnvVarUpdate(currentName, editor.ResultName, editor.ResultValue, target);
                        RefreshGridData(target);
                        ShowToast("✓  Updated!");
                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show(ErrorNoExit, $"Failed to update environment variable:\n{ex.Message}");
                    }
                    finally
                    {
                        this.UseWaitCursor = false;
                        Cursor.Current = Cursors.Default;
                    }
                }
            }
        }

        private void ApplyEnvVarUpdate(string oldName, string newName, string newValue, EnvironmentVariableTarget target)
        {
            bool nameChanged = !string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase);

            if (nameChanged)
            {
                // Remove old variable name before writing new variable name
                Environment.SetEnvironmentVariable(oldName, null, target);
            }

            Environment.SetEnvironmentVariable(newName, newValue, target);
        }

        private void ContextMenu_DeleteClick(object? sender, EventArgs e)
        {
            if (_contextTargetGrid == null || _contextTargetRow < 0) return;
            if (_contextTargetRow >= _contextTargetGrid.Rows.Count) return;

            string currentName = _contextTargetGrid.Rows[_contextTargetRow].Cells["colName"].Value?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(currentName) || currentName == "(none)") return;

            EnvironmentVariableTarget target = ResolveTargetFromGrid(_contextTargetGrid);

            // Require admin elevation for System scope deletions
            if (target == EnvironmentVariableTarget.Machine && !EnsureAdminElevation())
            {
                return;
            }

            // Immediately delete environment variable without confirmation prompts
            try
            {
                this.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;

                Environment.SetEnvironmentVariable(currentName, null, target);
                RefreshGridData(target);
                ShowToast("✓  Deleted!");
            }
            catch (Exception ex)
            {
                MsgBox.Show(ErrorNoExit, $"Failed to delete environment variable:\n{ex.Message}");
            }
            finally
            {
                this.UseWaitCursor = false;
                Cursor.Current = Cursors.Default;
            }
        }
        #endregion



        #region AddNewVariable
        private void HandleAddVariable(EnvironmentVariableTarget target)
        {
            // Require admin elevation for System scope additions
            if (target == EnvironmentVariableTarget.Machine && !EnsureAdminElevation())
            {
                return;
            }

            using (Form_EditEnvVar editor = new Form_EditEnvVar(
                EditMode.Add, string.Empty, string.Empty, target))
            {
                if (editor.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        this.UseWaitCursor = true;
                        Cursor.Current = Cursors.WaitCursor;

                        Environment.SetEnvironmentVariable(editor.ResultName, editor.ResultValue, target);
                        RefreshGridData(target);
                        ShowToast("✓  Added!");
                    }
                    catch (Exception ex)
                    {
                        MsgBox.Show(ErrorNoExit, $"Failed to add environment variable:\n{ex.Message}");
                    }
                    finally
                    {
                        this.UseWaitCursor = false;
                        Cursor.Current = Cursors.Default;
                    }
                }
            }
        }
        #endregion



        #region ExcelExport
        private void BtnExport_Click(object? sender, EventArgs e)
        {
            try
            {
                // Verify Excel template exists before proceeding
                if (!File.Exists(TemplateXlsmFilePath))
                {
                    MsgBox.Show(ErrorNoExit, $"Excel template file not found:\r\n{TemplateXlsmFilePath}");
                    return;
                }

                // Ensure target Excel output directory exists
                if (!Directory.Exists(XlsmOutputDirectory))
                {
                    Directory.CreateDirectory(XlsmOutputDirectory);
                }

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string filePath = Path.Combine(XlsmOutputDirectory, $"EnvVars_{MachineName}_{timestamp}.xlsm");
                FileInfo templateInfo = new FileInfo(TemplateXlsmFilePath);
                FileInfo outputInfo = new FileInfo(filePath);

                using (ExcelPackage package = new ExcelPackage(outputInfo, templateInfo))
                {
                    WriteEnvVarSheet(package, "User Variables", _userVars);
                    WriteEnvVarSheet(package, "System Variables", _systemVars);

                    package.Save();

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
                MsgBox.Show(ErrorNoExit, $"Failed to export Excel file:\r\n{ex.Message}");
            }
        }

        private void WriteEnvVarSheet(ExcelPackage package, string sheetName, List<EnvVarEntry> vars)
        {
            ExcelWorksheet ws = package.Workbook.Worksheets[sheetName];
            int lastTemplateRow = ws.Dimension.End.Row;

            // Populate machine name and export timestamp in template header
            ws.Cells[1, 2].Value = $"{ws.Cells[1, 2].Text}{MachineName}";
            ws.Cells[2, 2].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Write environment variable entries starting at row 4
            int dataStartRow = 4;
            int lastWriteRow = dataStartRow - 1;

            if (vars.Count == 0)
            {
                ws.Cells[dataStartRow, 1].Value = "(none)";
                ws.Cells[dataStartRow, 2].Value = "No environment variables detected.";
                lastWriteRow = dataStartRow;
            }
            else
            {
                for (int i = 0; i < vars.Count; i++)
                {
                    int row = dataStartRow + i;
                    ws.Cells[row, 1].Value = vars[i].Name;
                    ws.Cells[row, 2].Value = vars[i].Value;
                    lastWriteRow = row;
                }
            }

            // Clear borders and apply static background fill to unused template rows
            if (lastWriteRow < lastTemplateRow)
            {
                using (ExcelRange unusedRange = ws.Cells[lastWriteRow + 1, 1, lastTemplateRow, 2])
                {
                    unusedRange.Value = null;
                    unusedRange.Style.Border.Top.Style = ExcelBorderStyle.None;
                    unusedRange.Style.Border.Bottom.Style = ExcelBorderStyle.None;
                    unusedRange.Style.Border.Left.Style = ExcelBorderStyle.None;
                    unusedRange.Style.Border.Right.Style = ExcelBorderStyle.None;
                    unusedRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    unusedRange.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#505050"));
                }
            }
        }
        #endregion



        #region HelperMethods
        private void RefreshGridData(EnvironmentVariableTarget target)
        {
            if (target == EnvironmentVariableTarget.User)
            {
                _userVars = GetScopedEnvVars(EnvironmentVariableTarget.User);
            }
            else
            {
                _systemVars = GetScopedEnvVars(EnvironmentVariableTarget.Machine);
            }

            ApplySearchFilter();
        }

        private List<EnvVarEntry> GetScopedEnvVars(EnvironmentVariableTarget target)
        {
            List<EnvVarEntry> results = new();

            try
            {
                IDictionary rawVars = Environment.Environment_Master(target);

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
                MsgBox.Show(Warning, $"Failed to retrieve {target} environment variables:\n{ex.Message}");
            }

            return results;
        }

        private void ApplySearchFilter()
        {
            string query = _txtSearch?.Text?.Trim() ?? string.Empty;

            List<EnvVarEntry> filteredUser = FilterVarList(_userVars, query);
            List<EnvVarEntry> filteredSystem = FilterVarList(_systemVars, query);

            _userPanel.Text = string.IsNullOrEmpty(query)
                ? $"  User Environment Variables ({_userVars.Count})"
                : $"  User Environment Variables ({filteredUser.Count} / {_userVars.Count})";

            _systemPanel.Text = string.IsNullOrEmpty(query)
                ? $"  System Environment Variables ({_systemVars.Count})"
                : $"  System Environment Variables ({filteredSystem.Count} / {_systemVars.Count})";

            PopulateGrid(_userGrid, filteredUser, "User");
            PopulateGrid(_systemGrid, filteredSystem, "System");
        }

        private List<EnvVarEntry> FilterVarList(List<EnvVarEntry> source, string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return source;
            }

            return source
                .Where(v => v.Name.Contains(query, StringComparison.OrdinalIgnoreCase)
                         || v.Value.Contains(query, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private void PopulateGrid(DataGridView grid, List<EnvVarEntry> vars, string scope)
        {
            int previousScrollRow = grid.FirstDisplayedScrollingRowIndex;

            grid.SuspendLayout();
            try
            {
                grid.Rows.Clear();

                if (vars.Count == 0)
                {
                    int idx = grid.Rows.Add("(none)", $"No matching {scope} environment variables detected.");
                    grid.Rows[idx].DefaultCellStyle.ForeColor = _textSecondary;
                    grid.Rows[idx].DefaultCellStyle.Font = _gridEmptyFont;
                    return;
                }

                foreach (EnvVarEntry entry in vars)
                {
                    int rowIndex = grid.Rows.Add(entry.Name, entry.Value);
                    DataGridViewRow row = grid.Rows[rowIndex];
                    row.Cells[0].ToolTipText = CELL_TOOLTIP_TEXT;
                    row.Cells[1].ToolTipText = CELL_TOOLTIP_TEXT;
                }

                // Restore scroll offset when refreshing existing list
                if (previousScrollRow >= 0 && previousScrollRow < grid.Rows.Count)
                {
                    grid.FirstDisplayedScrollingRowIndex = previousScrollRow;
                }
            }
            finally
            {
                grid.ResumeLayout();
            }
        }

        private void LayoutPanels()
        {
            if (_userPanel == null || _systemPanel == null) return;

            int clientTop = this.TitleHeight > 0 ? this.TitleHeight : 35;
            int availableWidth = this.ClientSize.Width - (GRID_MARGIN * 2);
            int buttonBarHeight = 56;
            int availableHeight = this.ClientSize.Height - clientTop - (GRID_MARGIN * 3) - buttonBarHeight;
            int panelHeight = (availableHeight - GRID_MARGIN) / 2;

            // User panel at top
            _userPanel.Location = new Point(GRID_MARGIN, clientTop + GRID_MARGIN);
            _userPanel.Size = new Size(availableWidth, panelHeight);

            // System panel below User panel
            _systemPanel.Location = new Point(GRID_MARGIN, clientTop + GRID_MARGIN + panelHeight + GRID_MARGIN);
            _systemPanel.Size = new Size(availableWidth, panelHeight);

            // Position grids inside panels
            LayoutGridInPanel(_userGrid, _userPanel);
            LayoutGridInPanel(_systemGrid, _systemPanel);

            // Position Add buttons in top-right header of each panel
            LayoutAddButton(_btnAddUser, _userPanel);
            LayoutAddButton(_btnAddSystem, _systemPanel);

            // Position bottom bar controls (Search box, Clear button, Export button)
            int bottomBarY = this.ClientSize.Height - 42 - GRID_MARGIN;
            if (_txtSearch != null)
            {
                _txtSearch.Location = new Point(GRID_MARGIN, bottomBarY + 2);
            }

            if (_btnClearSearch != null && _txtSearch != null)
            {
                _btnClearSearch.Location = new Point(_txtSearch.Right + 8, bottomBarY + 2);
            }

            if (_btnExport != null)
            {
                _btnExport.Location = new Point(
                    this.ClientSize.Width - _btnExport.Width - GRID_MARGIN,
                    bottomBarY);
            }

            // Position floating toast notification above Export button
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

        private EnvironmentVariableTarget ResolveTargetFromGrid(DataGridView grid)
        {
            return grid == _userGrid
                ? EnvironmentVariableTarget.User
                : EnvironmentVariableTarget.Machine;
        }

        private bool EnsureAdminElevation()
        {
            if (IsRunningAsAdmin())
            {
                return true;
            }

            DialogResult result = MsgBox.Show(Question,
                "Modifying System environment variables requires administrator privileges.\r\n\r\n" +
                "The application will relaunch with elevated permissions.\r\nAny unsaved work will be preserved in Windows.",
                "Admin Privileges Required");
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
                return false;
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // User declined UAC elevation prompt
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

        private void ShowToast(string message)
        {
            _toastTimer.Stop();
            _toastLabel.Text = message;
            _toastLabel.Visible = true;
            _toastLabel.BringToFront();
            _toastTimer.Start();
        }
        #endregion
    }



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
        const int STANDARD_DIALOG_WIDTH = 480;
        const int STANDARD_DIALOG_HEIGHT = 290;
        const int PATH_DIALOG_WIDTH = 680;
        const int PATH_DIALOG_HEIGHT = 520;

        private static readonly char[] _invalidNameChars = { '=', '%', '<', '>', '^', '&', '|', ':' };

        private EditMode _mode { get; }
        private string _originalName { get; }
        private string _originalValue { get; }
        private EnvironmentVariableTarget _target { get; }
        private bool _isPathVar { get; }

        private UITextBox _txtName { get; set; }
        private UITextBox _txtValue { get; set; }
        private UIPanel _rtbBorderPanel { get; set; }
        private RichTextBox _rtbPathValue { get; set; }
        private UIButton _btnAction { get; set; }
        private UIButton _btnCancel { get; set; }
        private Label _lblError { get; set; }

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
            _isPathVar = string.Equals(currentName, "Path", StringComparison.OrdinalIgnoreCase);

            InitializeDialog();
            InitializeControls();
        }

        private void InitializeDialog()
        {
            string scopeLabel = _target == EnvironmentVariableTarget.User ? "User" : "System";
            string actionLabel = _mode == EditMode.Update ? "Update" : "Add";
            string pathSuffix = _isPathVar ? " (Path List Editor)" : string.Empty;

            int dialogWidth = _isPathVar ? PATH_DIALOG_WIDTH : STANDARD_DIALOG_WIDTH;
            int dialogHeight = _isPathVar ? PATH_DIALOG_HEIGHT : STANDARD_DIALOG_HEIGHT;

            this.Text = $"{actionLabel} {scopeLabel} Environment Variable{pathSuffix}";
            this.Size = new Size(dialogWidth, dialogHeight);
            this.MinimumSize = _isPathVar ? new Size(560, 400) : this.Size;
            this.MaximumSize = _isPathVar ? new Size(1200, 900) : new Size(640, STANDARD_DIALOG_HEIGHT);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = _isPathVar ? FormBorderStyle.Sizable : FormBorderStyle.FixedDialog;
            this.MaximizeBox = _isPathVar;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.KeyPreview = true;

            // Sunny.UI dark theming
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

            // Allow Escape key to cancel dialog cleanly
            this.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    this.DialogResult = DialogResult.Cancel;
                    this.Close();
                }
            };
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
                RectColor = Color.FromArgb(58, 60, 82),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(_txtName);
            currentY += 40;

            // Value label
            UILabel lblValue = new UILabel
            {
                Text = _isPathVar ? "Value (one path per line — semicolons normalized automatically):" : "Value:",
                Location = new Point(labelX, currentY),
                Size = new Size(fieldWidth, 20),
                Font = new Font("Segoe UI Semibold", 9.5f),
                ForeColor = accentColor,
                Style = UIStyle.Custom,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblValue);
            currentY += 24;

            // Action button metrics calculated from bottom edge
            string actionText = _mode == EditMode.Update ? "Update" : "Add";
            int btnWidth = 94;
            int btnHeight = 34;
            int btnSpacing = 10;
            int bottomMargin = 18;
            int btnY = this.ClientSize.Height - btnHeight - bottomMargin;
            int btnRightEdge = fieldX + fieldWidth;

            if (_isPathVar)
            {
                // Multi-line RichTextBox at bottom for Path / PATH variable values
                int rtbHeight = btnY - currentY - 14;

                _rtbBorderPanel = new UIPanel
                {
                    Location = new Point(fieldX, currentY),
                    Size = new Size(fieldWidth, rtbHeight),
                    Style = UIStyle.Custom,
                    FillColor = Color.FromArgb(40, 42, 60),
                    RectColor = Color.FromArgb(58, 60, 82),
                    Radius = 4,
                    RadiusSides = UICornerRadiusSides.All,
                    Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
                };

                _rtbPathValue = new RichTextBox
                {
                    Location = new Point(6, 6),
                    Size = new Size(fieldWidth - 12, rtbHeight - 12),
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.FromArgb(40, 42, 60),
                    ForeColor = Color.FromArgb(230, 230, 240),
                    Font = new Font("Cascadia Code", 9.5f),
                    WordWrap = false,
                    DetectUrls = false,
                    ScrollBars = RichTextBoxScrollBars.Both,
                    Text = NormalizePathForDisplay(_originalValue),
                    Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
                };

                _rtbBorderPanel.Controls.Add(_rtbPathValue);
                this.Controls.Add(_rtbBorderPanel);
            }
            else
            {
                // Standard single-line UITextBox for general environment variables
                _txtValue = new UITextBox
                {
                    Text = _originalValue,
                    Location = new Point(fieldX, currentY),
                    Size = new Size(fieldWidth, 32),
                    Font = new Font("Cascadia Code", 10f),
                    Style = UIStyle.Custom,
                    FillColor = Color.FromArgb(40, 42, 60),
                    ForeColor = Color.FromArgb(230, 230, 240),
                    RectColor = Color.FromArgb(58, 60, 82),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                this.Controls.Add(_txtValue);
            }

            // Validation error label
            _lblError = new Label
            {
                AutoSize = false,
                Size = new Size(fieldWidth - (btnWidth * 2) - btnSpacing - 12, 24),
                Location = new Point(fieldX, btnY + 6),
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(255, 85, 85),
                BackColor = Color.Transparent,
                Text = string.Empty,
                Visible = false,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(_lblError);

            // Cancel button
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
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            _btnCancel.Click += (s, e) =>
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            };
            this.Controls.Add(_btnCancel);

            // Primary action button (Update or Add)
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
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            _btnAction.Click += BtnAction_Click;
            this.Controls.Add(_btnAction);
        }

        private string NormalizePathForDisplay(string rawValue)
        {
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return string.Empty;
            }

            // Replace every semicolon delimiter with Environment.NewLine for line-by-line editing
            return rawValue
                .TrimEnd(';')
                .Replace(";", Environment.NewLine);
        }
        #endregion



        #region Validation
        private void BtnAction_Click(object? sender, EventArgs e)
        {
            string name = _txtName.Text?.Trim() ?? string.Empty;
            string value = _isPathVar
                ? CollapsePathFromDisplay(_rtbPathValue.Text)
                : (_txtValue.Text?.Trim() ?? string.Empty);

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
                if (_isPathVar) _rtbPathValue.Focus();
                else _txtValue.Focus();
                return;
            }

            // Validate name does not contain invalid characters
            foreach (char c in _invalidNameChars)
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
                if (_isPathVar) _rtbPathValue.Focus();
                else _txtValue.Focus();
                return;
            }

            // Validate name does not start with a digit
            if (char.IsDigit(name[0]))
            {
                ShowError("Name cannot start with a digit.");
                _txtName.Focus();
                return;
            }

            bool nameChanged = !string.Equals(_originalName, name, StringComparison.OrdinalIgnoreCase);

            // Update mode: verify changes exist
            if (_mode == EditMode.Update)
            {
                string baselineValue = _isPathVar ? CollapsePathFromDisplay(_originalValue) : _originalValue;
                bool valueChanged = !string.Equals(baselineValue, value, StringComparison.Ordinal);

                if (!nameChanged && !valueChanged)
                {
                    ShowError("No changes detected.");
                    return;
                }
            }

            // Prevent overwriting an existing variable when adding or renaming
            if (_mode == EditMode.Add || (_mode == EditMode.Update && nameChanged))
            {
                string? existing = Environment.GetEnvironmentVariable(name, _target);
                if (existing != null)
                {
                    ShowError($"'{name}' already exists in this scope.");
                    _txtName.Focus();
                    return;
                }
            }

            // Commit validated result
            ResultName = name;
            ResultValue = value;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private string CollapsePathFromDisplay(string multilineValue)
        {
            if (string.IsNullOrWhiteSpace(multilineValue))
            {
                return string.Empty;
            }

            // Collapse newline-separated paths back into a clean semicolon-delimited string
            string[] segments = multilineValue
                .Split(new[] { Environment.NewLine, "\n", "\r", ";" }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            return string.Join(";", segments);
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
        #region RendererConfig
        private Color _borderColor { get; }
        private Color _hoverColor { get; }
        #endregion



        #region RendererConstructor
        public DarkMenuRenderer(Color borderColor, Color hoverColor)
            : base(new DarkColorTable(borderColor))
        {
            _borderColor = borderColor;
            _hoverColor = hoverColor;
        }
        #endregion



        #region PaintOverrides
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
        #endregion
    }

    public class DarkColorTable : ProfessionalColorTable
    {
        #region ColorTableProperties
        private Color _borderColor { get; }

        public override Color MenuBorder => _borderColor;
        public override Color MenuItemBorder => Color.Transparent;
        public override Color ToolStripDropDownBackground => Color.FromArgb(40, 42, 60);
        public override Color ImageMarginGradientBegin => Color.FromArgb(40, 42, 60);
        public override Color ImageMarginGradientMiddle => Color.FromArgb(40, 42, 60);
        public override Color ImageMarginGradientEnd => Color.FromArgb(40, 42, 60);
        public override Color SeparatorDark => Color.FromArgb(58, 60, 82);
        public override Color SeparatorLight => Color.FromArgb(58, 60, 82);
        #endregion



        #region ColorTableConstructor
        public DarkColorTable(Color borderColor)
        {
            _borderColor = borderColor;
        }
        #endregion
    }
    #endregion



    #region EnvVarEntryModel
    public class EnvVarEntry
    {
        #region ModelProperties
        public string Name { get; }
        public string Value { get; }
        #endregion



        #region ModelConstructor
        public EnvVarEntry(string name, string value)
        {
            Name = name;
            Value = value;
        }
        #endregion
    }
    #endregion
}
