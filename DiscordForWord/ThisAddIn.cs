using DiscordRPC;
using Microsoft.Office.Interop.Word;
using System;

namespace DiscordForWord
{
    public partial class ThisAddIn
    {
        public DiscordRpcClient client;
        private static RichPresence presence = Shared.Shared.getNewPresence("word");

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            client = new DiscordRpcClient(Shared.Shared.getString("discordID"));
            client.Initialize();
            client.SetPresence(presence);

            this.Application.WindowDeactivate += new ApplicationEvents4_WindowDeactivateEventHandler(
                Application_WindowDeactivate);
            this.Application.WindowActivate += new ApplicationEvents4_WindowActivateEventHandler(
                Application_WindowActivate);
            this.Application.DocumentOpen += new ApplicationEvents4_DocumentOpenEventHandler(
                Application_DocumentOpen);
            ((ApplicationEvents4_Event)this.Application).NewDocument += new ApplicationEvents4_NewDocumentEventHandler(
                Application_DocumentOpen);
            this.Application.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(
                Application_WindowSelectionChange);
            this.Application.DocumentChange += new ApplicationEvents4_DocumentChangeEventHandler(
                Application_DocumentChange);

            try
            {
                // Use the currently opened document
                Document doc = this.Application.ActiveDocument;
                Application_DocumentOpen(doc);
            }
            catch
            {
                // Use the default presence when there is no current document

            }
        }

        private void Application_DocumentChange()
        {
            if (Application.Documents.Count == 1)
            {
                Application_WindowSelectionChange(Application.Selection);
            }
        }

        private void Application_WindowDeactivate(Document doc, Window wn)
        {
            if (Application.Documents.Count == 1)
            {
                presence.Details = Shared.Shared.getString("tabOut");
                presence.State = null;
                presence.Party = null;
                presence.Assets.LargeImageKey = "word_nothing";
            }

            client.SetPresence(presence);
        }

        private void Application_WindowClose()
        {
            if (Application.Documents.Count > 1)
            {
                presence.Details = "" + Application.Documents.Count;
                Application_WindowSelectionChange(Application.Selection);
            }
            else
            {
                presence.Details = Shared.Shared.getString("tabOut") + Application.Documents.Count;
                presence.State = null;
                presence.Party = null;
                presence.Assets.LargeImageKey = "word_nothing";
            }

            client.SetPresence(presence);
        }

        private void Application_DocumentOpen(Document doc)
        {
            Application_WindowSelectionChange(Application.Selection);

            ((DocumentEvents2_Event)doc).Close += new DocumentEvents2_CloseEventHandler(Application_WindowClose);
        }

        private void Application_WindowActivate(Document doc, Window wn)
        {
            Application_WindowSelectionChange(Application.Selection);
        }

        public void Application_WindowSelectionChange(Selection sel)
        {
            Range range = Application.ActiveDocument.Content;

            presence.Details = Application.ActiveDocument.Name;
            presence.State = Shared.Shared.getString("editingPage");
            presence.Assets.LargeImageKey = "word_editing";
            presence.Party = new Party()
            {
                ID = Secrets.CreateFriendlySecret(new Random()),
                Max = range.ComputeStatistics(WdStatistic.wdStatisticPages),
                Size = (int)sel.get_Information(WdInformation.wdActiveEndPageNumber)
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
