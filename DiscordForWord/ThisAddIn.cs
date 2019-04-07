using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using DiscordRPC;
using Shared;
using Microsoft.Office.Interop.Word;

namespace DiscordForWord
{
    public partial class ThisAddIn
    {
        public DiscordRpcClient client;
        private static RichPresence presence = Shared.Shared.getNewPresence("word");

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            client = new DiscordRpcClient(Shared.Shared.getString("discordID"), true, -1);
            client.Initialize();
            client.SetPresence(presence);

            this.Application.WindowDeactivate += new ApplicationEvents4_WindowDeactivateEventHandler(
                Application_WindowDeactivate);
            this.Application.WindowActivate += new ApplicationEvents4_WindowActivateEventHandler(
                Application_WindowActivate);
            this.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(
                Application_DocumentOpen);
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(
                Application_DocumentOpen);

            try
            {
                // Use the currently opened document
                Word.Document doc = this.Application.ActiveDocument;
                Application_DocumentOpen(doc);
            } catch
            {
                // Use the default presence when there is no current document
                
            }
        }

        private void Application_DocumentOpen(Word.Document doc)
        {
            presence.Details = Application.ActiveDocument.Name;
            presence.State = Shared.Shared.getString("editing");
            presence.Assets.LargeImageKey = "word_editing";

            client.SetPresence(presence);
        }

        private void Application_WindowDeactivate(Word.Document doc, Window wn)
        {
            if (Application.Documents.Count == 1)
            {
                presence.Details = Shared.Shared.getString("tabOut");
                presence.State = null;
                presence.Assets.LargeImageKey = "word_nothing";
            }
            else
            {
                presence.Details = Application.ActiveDocument.Name;
            }

            client.SetPresence(presence);
        }

        private void Application_WindowActivate(Word.Document doc, Window wn)
        {
            Application_DocumentOpen(doc);
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
