using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using DiscordRPC;
using DiscordRPC.Logging;
using DiscordRPC.Message;
using System.Threading;

namespace DiscordForPowerPoint
{
    public partial class ThisAddIn
    {
        public DiscordRpcClient client;
        private static int DiscordPipe = -1;
        private static string ClientID = "470239659591598091";
        private static LogLevel DiscordLogLevel = LogLevel.Info;

        private static RichPresence presence = new RichPresence()
        {
            Details = "Untitled Presentation",
            State = "Editing",
            Assets = new Assets()
            {
                LargeImageKey = "welcome",
                LargeImageText = "Not Editing",
                SmallImageKey = "powerpoint"
            }
        };

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            client = new DiscordRpcClient(ClientID, true, DiscordPipe);
            client.Logger = new DiscordRPC.Logging.ConsoleLogger() { Level = DiscordLogLevel, Coloured = true };

            client.Initialize();

            client.SetPresence(presence);

            Debug.Print("aaa");
            this.Application.PresentationNewSlide += 
                new PowerPoint.EApplication_PresentationNewSlideEventHandler(
                Application_PresentationNewSlide);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            client.Dispose();
        }

        private void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
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
