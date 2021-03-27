using System;
using DiscordRPC;
using Microsoft.Office.Interop.Outlook;

namespace DiscordForOutlook
{
    public partial class ThisAddIn
    {
        public DiscordRpcClient client;
        private static RichPresence presence = Shared.Shared.getNewPresence("outlook");

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            client = new DiscordRpcClient(Shared.Shared.getString("discordID"));
            client.Initialize();
            presence.State = null;
            presence.Details = null;
            presence.Assets.LargeImageKey = "outlook_info";
            client.SetPresence(presence);

            ((ApplicationEvents_11_Event)Application).Quit += new ApplicationEvents_11_QuitEventHandler(ThisAddIn_Quit);
        }

        private void ThisAddIn_Quit()
        {
            client.Dispose();
            return;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
