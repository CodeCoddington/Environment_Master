namespace Environment_Master
{
    internal static class Program
    {
        #region MainExecution
        [STAThread]
        static void Main()
        {
            // Initialize app configuration and launch main form
            ApplicationConfiguration.Initialize();
            Application.Run(new Form_EnvVars());
        }
        #endregion
    }
}