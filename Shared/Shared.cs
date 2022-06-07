using DiscordRPC;
using System.Collections.Generic;

namespace Shared
{
    public class Shared
    {
        private static IDictionary<string, string> Strings = new Dictionary<string, string>()
        {
            {"discordID", "470239659591598091" },
            {"noFile", "No File Open" },
            {"tabOut", "Not Active"},
            {"welcome", "Welcome Screen" },
            {"editing", "Editing File" },
            {"editingSlide", "Editing Slide" },
            {"editingPage", "Editing Page" },
            {"presenting", "Presenting" },
            {"unknown_key", "[Unknown]" }
        };

        public static string getString(string key)
        {
            if (Strings.ContainsKey(key))
            {
                return Strings[key];
            }
            else
            {
                return getString("unknown_key");
            }
        }

        public static bool isEnabled()
        {
            return Options.Default.enabled;
        }

        public static RichPresence getNewPresence(Program program)
        {
            // Lowercase the enum name
            string programName = program.ToString().ToLower();

            // Create a new Rich Presence for the specific Office program
            return new RichPresence()
            {
                Details = getString("noFile"),
                State = getString("welcome"),
                Assets = new Assets()
                {
                    LargeImageKey = programName + "_welcome",
                    LargeImageText = Products.getProductName(program),
                    SmallImageKey = programName
                }
            };
        }
    }
}
