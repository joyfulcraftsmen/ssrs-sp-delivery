using System;
using System.Collections;

using Microsoft.ReportingServices.Interfaces;

namespace JC.SSRS.Extensions.SharePointDelivery
{
    internal class SubscriptionData
    {

        // Initalize variables to default values
        public string Server = "";
        public string path = "";
        public string file = "";
        public string renderingFormat = "PDF";

        // Initialize setting names
        internal const string SERVER = "Server";
        internal const string PATH = "Path";
        internal const string FILE = "File";
        internal const string RENDERINGFORMAT = "RenderingFormat";

        public SubscriptionData()
        {
            // TODO: Add constructor code here
        }

        // Populate the object from an array of setting elements
        // No validation is done, it is assumed that the settings
        // contains all relevant information
        public void FromSettings(Setting[] settings)
        {
            foreach (Setting setting in settings)
            {
                switch (setting.Name)
                {
                    case (SERVER):
                        Server = setting.Value;
                        break;
                    case (PATH):
                        this.path = setting.Value;
                        break;
                    case (FILE):
                        this.file = setting.Value;
                        break;
                    case (RENDERINGFORMAT):
                        this.renderingFormat = setting.Value;
                        break;
                    default:
                        break;
                }
            }
        }

        // Creates an array of the settings
        public Setting[] ToSettingArray()
        {
            ArrayList list = new ArrayList();

            list.Add(CreateSetting(SERVER, this.Server));
            list.Add(CreateSetting(PATH, this.path));
            list.Add(CreateSetting(FILE, this.file));
            list.Add(CreateSetting(RENDERINGFORMAT, this.renderingFormat));

            return list.ToArray(typeof(Setting)) as Setting[];
        }

        // Creates a single instance of a Setting
        private static Setting CreateSetting(string name, string val)
        {
            Setting setting = new Setting();
            setting.Name = name;
            setting.Value = val;

            return setting;
        }
    }
}


