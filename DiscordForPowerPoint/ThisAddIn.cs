using System;
using DiscordRPC;
using Microsoft.Office.Interop.PowerPoint;

namespace DiscordForPowerPoint
{
    public partial class ThisAddIn
    {
        public DiscordRpcClient client;
        private static RichPresence presence = Shared.Shared.getNewPresence("powerpoint");

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            client = new DiscordRpcClient(Shared.Shared.getString("discordID"));
            client.Initialize();
            client.SetPresence(presence);

            // An event handler for when a new slide is created
            this.Application.PresentationNewSlide += 
                new EApplication_PresentationNewSlideEventHandler(
                Application_PresentationNewSlide);

            // An event handler for any time a slide / slides / inbetween slides is selected
            this.Application.SlideSelectionChanged +=
                new EApplication_SlideSelectionChangedEventHandler(
                Application_SlideSelectionChanged);

            // An event handler for when a file is closed.
            // Final = Actually closed
            this.Application.PresentationCloseFinal +=
                new EApplication_PresentationCloseFinalEventHandler(
                Application_PresentationCloseFinal);

            // Event handlers for when a file is created, opened, saved, or slide show ends.
            this.Application.AfterNewPresentation +=
                new EApplication_AfterNewPresentationEventHandler(
                Application_AfterPresentationOpenEvent);
            this.Application.AfterPresentationOpen +=
                new EApplication_AfterPresentationOpenEventHandler(
                Application_AfterPresentationOpenEvent);
            this.Application.PresentationSave +=
                new EApplication_PresentationSaveEventHandler(
                Application_AfterPresentationOpenEvent);
            this.Application.SlideShowEnd +=
                new EApplication_SlideShowEndEventHandler(
                Application_AfterPresentationOpenEvent);

            // An event handler for when a slide show starts, or goes onto a new slide
            this.Application.SlideShowNextSlide +=
                new EApplication_SlideShowNextSlideEventHandler(
                Application_SlideShowNextSlide);
        }

        // When Microsoft PowerPoint shuts down, delete the RPC client.
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            client.Dispose();
        }

        private void Application_PresentationNewSlide(Slide Sld)
        {
            // Assumption: People start on slide 1 when creating a file
            presence.Party = new Party()
            {
                Max = Application.ActivePresentation.Slides.Count,
                Size = 1
            };

            client.SetPresence(presence);
        }

        private void Application_SlideSelectionChanged(SlideRange SldRange)
        {
            if (SldRange.Count > 0)
            {
                presence.Details = SldRange.Application.ActivePresentation.Name;
                presence.State = Shared.Shared.getString("editing");
                presence.Assets.LargeImageKey = "editing";
                presence.Party = new Party()
                {
                    ID = Secrets.CreateFriendlySecret(new Random()),
                    Max = Application.ActivePresentation.Slides.Count,
                    Size = SldRange[1].SlideNumber
                };
                client.SetPresence(presence);
            }
        }

        public void Application_PresentationCloseFinal(Presentation Pres)
        {
            // There's only one presentation left - the current one
            if (Application.Presentations.Count == 1)
            {
                presence.Details = Shared.Shared.getString("noFile");
                presence.State = null;
                presence.Party = null;
                presence.Assets.LargeImageKey = "nothing";
            } else
            {
                presence.Details = Application.ActivePresentation.Name;
            }

            client.SetPresence(presence);
        }

        public void Application_AfterPresentationOpenEvent(Presentation Pres)
        {
            presence.Details = Pres.Name;
            presence.State = Shared.Shared.getString("editingSlide");
            presence.Assets.LargeImageKey = "editing";

            // Slide selection is also triggered - Don't need to set presence
        }

        public void Application_SlideShowNextSlide(SlideShowWindow Wn)
        {
            presence.Details = Wn.Presentation.Name;
            presence.State = Shared.Shared.getString("presenting");
            presence.Assets.LargeImageKey = "present";
            presence.Party = new Party()
            {
                ID = Secrets.CreateFriendlySecret(new Random()),
                Max = Wn.Presentation.Slides.Count,
                Size = Wn.View.CurrentShowPosition
            };
            client.SetPresence(presence);
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
