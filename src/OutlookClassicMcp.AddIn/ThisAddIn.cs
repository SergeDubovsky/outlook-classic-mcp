using System;
using System.Diagnostics;
using OutlookClassicMcp.AddIn.Runtime;

namespace OutlookClassicMcp.AddIn
{
    public partial class ThisAddIn
    {
        private AddInHost? _host;
#if OUTLOOK_MCP_SMOKE_SEEDER
        private Phase4FixtureSeeder? _fixtureSeeder;
#endif

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var started = Stopwatch.GetTimestamp();
            try
            {
#if OUTLOOK_MCP_SMOKE_SEEDER
                if (Phase4FixtureSeeder.TryStart(Application, out var fixtureSeeder))
                {
                    _fixtureSeeder = fixtureSeeder;
                    return;
                }
#endif
                var host = new AddInHost(Application);
                _host = host;
                host.Start(started);
            }
            catch
            {
                // Startup failures must not destabilize Outlook.
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
#if OUTLOOK_MCP_SMOKE_SEEDER
            _fixtureSeeder?.Dispose();
#endif
            _ = _host?.BeginShutdown();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
