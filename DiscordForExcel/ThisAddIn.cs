using DiscordRPC;
using Microsoft.Office.Interop.Excel;
using System;

namespace DiscordForExcel
{
    public partial class ThisAddIn
    {
        public DiscordRpcClient client;
        private static RichPresence presence = Shared.Shared.getNewPresence("excel");

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            client = new DiscordRpcClient(Shared.Shared.getString("discordID"));
            client.Initialize();
            client.SetPresence(presence);

            this.Application.WorkbookDeactivate += new AppEvents_WorkbookDeactivateEventHandler(
                Application_WorkbookDeactivate);
            this.Application.WorkbookOpen += new AppEvents_WorkbookOpenEventHandler(
                Application_WorkbookOpen);
            ((AppEvents_Event)this.Application).NewWorkbook += new AppEvents_NewWorkbookEventHandler(
                Application_WorkbookOpen);
        }

        private void Application_WorkbookOpen(Workbook Wb)
        {
            presence.Details = Application.ActiveWorkbook.Name;
            presence.State = Shared.Shared.getString("editing");
            presence.Assets.LargeImageKey = "excel_editing";

            client.SetPresence(presence);
        }

        private void Application_WorkbookDeactivate(Workbook Wb)
        {
            if (Application.Workbooks.Count == 1)
            {
                presence.Details = Shared.Shared.getString("noFile");
                presence.State = null;
                presence.Assets.LargeImageKey = "excel_nothing";
            }
            else
            {
                presence.Details = Application.ActiveWorkbook.Name;
            }

            client.SetPresence(presence);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            client.Dispose();
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
