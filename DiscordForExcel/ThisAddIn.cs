using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using DiscordRPC;
using DiscordRPC.Logging;
using Microsoft.Office.Interop.Excel;

namespace DiscordForExcel
{
    public partial class ThisAddIn
    {
        private static IDictionary<int, string> OfficeVersions = new Dictionary<int, string>() {
            {6, "4.x"},
            {7, "95"},
            {8, "97"},
            {9, "2000"},
            {10, "XP"},
            {11, "2003"},
            {12, "2007"},
            {14, "2010"},
            {15, "2013"},
            {16, "2016"},
            {17, "2017"}
        };

        public DiscordRpcClient client;
        private static int DiscordPipe = -1;
        private static string ClientID = "470239659591598091";
        private static LogLevel DiscordLogLevel = LogLevel.Info;

        private static RichPresence presence = new RichPresence()
        {
            Details = "No File Open",
            State = "Welcome Screen",
            Assets = new Assets()
            {
                LargeImageKey = "excel_welcome",
                LargeImageText = "Microsoft Excel " + OfficeVersions[Process.GetCurrentProcess().MainModule.FileVersionInfo.ProductMajorPart],
                SmallImageKey = "excel"
            }
        };

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            client = new DiscordRpcClient(ClientID, true, DiscordPipe);
            client.Logger = new DiscordRPC.Logging.ConsoleLogger() { Level = DiscordLogLevel, Coloured = true };
            client.Initialize();
            client.SetPresence(presence);

            this.Application.WorkbookDeactivate += new AppEvents_WorkbookDeactivateEventHandler(
                Application_WorkbookDeactivate);
            this.Application.WorkbookOpen += new AppEvents_WorkbookOpenEventHandler(
                Application_WorkbookOpen);
            ((Excel.AppEvents_Event)this.Application).NewWorkbook += new AppEvents_NewWorkbookEventHandler(
                Application_WorkbookOpen);
        }

        private void Application_WorkbookOpen(Workbook Wb)
        {
            presence.Details = Wb.Name;
            presence.State = "Editing";
            presence.Assets.LargeImageKey = "excel_editing";

            client.SetPresence(presence);
        }

        private void Application_WorkbookDeactivate(Workbook Wb)
        {
            presence = new RichPresence()
            {
                Details = "No File Open",
                State = "No File Open",
                Assets = new Assets()
                {
                    LargeImageKey = "excel_nothing",
                    LargeImageText = "Microsoft Excel " + OfficeVersions[Process.GetCurrentProcess().MainModule.FileVersionInfo.ProductMajorPart],
                    SmallImageKey = "excel"
                }
            };

            client.SetPresence(presence);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
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
