﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using DiscordRPC;
using DiscordRPC.Logging;
using Microsoft.Office.Interop.PowerPoint;

namespace DiscordForPowerPoint
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
                LargeImageKey = "welcome",
                SmallImageKey = "powerpoint"
            }
        };

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            client = new DiscordRpcClient(ClientID, true, DiscordPipe);
            client.Logger = new DiscordRPC.Logging.ConsoleLogger() { Level = DiscordLogLevel, Coloured = true };

            presence.Assets.LargeImageText = "Microsoft PowerPoint " + OfficeVersions[Process.GetCurrentProcess().MainModule.FileVersionInfo.ProductMajorPart];

            client.Initialize();

            client.SetPresence(presence);

            // An event handler for when a new slide is created
            this.Application.PresentationNewSlide += 
                new PowerPoint.EApplication_PresentationNewSlideEventHandler(
                Application_PresentationNewSlide);

            // An event handler for any time a slide / slides / inbetween slides is selected
            this.Application.SlideSelectionChanged +=
                new PowerPoint.EApplication_SlideSelectionChangedEventHandler(
                Application_SlideSelectionChanged);

            // An event handler for when a file is closed.
            // Final = Actually closed
            this.Application.PresentationCloseFinal +=
                new PowerPoint.EApplication_PresentationCloseFinalEventHandler(
                Application_PresentationCloseFinal);

            // Event handlers for when a file is created, opened, saved, or slide show ends.
            this.Application.AfterNewPresentation +=
                new PowerPoint.EApplication_AfterNewPresentationEventHandler(
                Application_AfterPresentationOpenEvent);
            this.Application.AfterPresentationOpen +=
                new PowerPoint.EApplication_AfterPresentationOpenEventHandler(
                Application_AfterPresentationOpenEvent);
            this.Application.PresentationSave +=
                new PowerPoint.EApplication_PresentationSaveEventHandler(
                Application_AfterPresentationOpenEvent);
            this.Application.SlideShowEnd +=
                new PowerPoint.EApplication_SlideShowEndEventHandler(
                Application_AfterPresentationOpenEvent);

            // An event handler for when a slide show starts, or goes onto a new slide
            this.Application.SlideShowNextSlide +=
                new PowerPoint.EApplication_SlideShowNextSlideEventHandler(
                Application_SlideShowNextSlide);
        }

        // When Microsoft PowerPoint shuts down, delete the RPC client.
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
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
                presence.State = "Editing";
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
            if (Application.Presentations.Count == 1)
            {
                Debug.Print("Presentation Closed");
                presence.Details = "No File Open";
                presence.State = "No File Open";
                presence.Party = null;
                presence.Assets.LargeImageKey = "nothing";
            }
            else
            {
                presence.Details = Application.ActivePresentation.Name;
            }

            client.SetPresence(presence);
        }

        public void Application_AfterPresentationOpenEvent(Presentation Pres)
        {
            presence.Details = Pres.Name;
            presence.State = "Editing";
            presence.Assets.LargeImageKey = "editing";

            // Slide selection is also triggered - Don't need to set presence
        }

        public void Application_SlideShowNextSlide(SlideShowWindow Wn)
        {
            presence.Details = Wn.Presentation.Name;
            presence.State = "Presenting";
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
