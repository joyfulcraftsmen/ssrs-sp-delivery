using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml.Linq;
using Microsoft.ReportingServices.Interfaces;
using cli = Microsoft.SharePoint.Client;

namespace JC.SSRS.Extensions.SharePointDelivery
{
    public class SharePointDeliveryExtension : IDeliveryExtension, IExtension
    {
        private Setting[] m_settings { get; set; } = null;
        public Setting[] ExtensionSettings
        {
            get
            {
                if (this.m_settings == null)
                {
                    m_settings = new Setting[4];
                    m_settings[0] = new Setting()
                    {
                        Name = "Server",
                        ReadOnly = false,
                        Required = true
                    };
                    m_settings[1] = new Setting()
                    {
                        Name = "Path",
                        ReadOnly = false,
                        Required = true
                    };
                    m_settings[2] = new Setting()
                    {
                        Name = "File",
                        ReadOnly = false,
                        Required = false
                    };
                    m_settings[3] = new Setting()
                    {
                        Name = "RenderingFormat",
                        ReadOnly = false,
                        Required = false
                    };
                }
                return m_settings;
            }
        }

        private bool isPrivilegedUser = false;

        public bool IsPrivilegedUser
        {
            set
            {
                this.isPrivilegedUser = value;
            }
        }

        public string LocalizedName
        {
            get
            {
                return "SharePointLibrary";
            }
        }

        private IDeliveryReportServerInformation reportServerInformation = null;

        public IDeliveryReportServerInformation ReportServerInformation
        {
            set
            {
                this.reportServerInformation = value;
            }
        }

        private List<string> validServers { get; set; } = new List<string>();

        public bool Deliver(Notification notification)
        {
            try
            {
                bool success = false;
                notification.Status = "Processing";

                try
                {
                    Setting[] userSettings = notification.UserData;
                    SubscriptionData subscriptionData = new SubscriptionData();
                    subscriptionData.FromSettings(userSettings);

                    // do the job
                    string server = userSettings.First(s => s.Name == "Server").Value;
                    string path = userSettings.First(s => s.Name == "Path").Value;
                    var fileSetting = userSettings.FirstOrDefault(s => s.Name == "File");
                    string file = fileSetting?.Value ?? notification.Report.Name + ".pdf";
                    var renderingFormatSetting = userSettings.FirstOrDefault(s => s.Name == "RenderingFormat");
                    string renderingFormat = renderingFormatSetting?.Value ?? "PDF";

                    if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(path) || string.IsNullOrEmpty(file))
                        throw new Exception("Server name, path and file name are required parameters. At least one of them is missing.");
                    
                    var context = new cli.ClientContext(server);
                    context.Credentials = CredentialCache.DefaultCredentials;
                    var m_files = notification.Report.Render(renderingFormat, "<DeviceInfo></DeviceInfo>");

                    // get full path
                    string fullPath = path.Replace(@"\", @"/");
                    if (!fullPath.EndsWith("/"))
                        fullPath += @"/";
                    if (!fullPath.StartsWith("/"))
                        fullPath = @"/" + fullPath;
                    fullPath += file;

                    using (Stream fs = m_files[0].Data)
                    {
                        
                        fs.Position = 0;
                        cli.File.SaveBinaryDirect(context, fullPath, fs, true);
                    }
                    notification.Status = "Success";
                    success = true;
                }
                catch(Exception ex)
                {
                    notification.Status = "Error: " + ex.Message;
                    success = false;
                }
                finally
                {
                    notification.Save();
                }

                return success;
            }
            catch 
            {
                return false;
            }
        }

        string configuration = null;

        public void SetConfiguration(string configuration)
        {
            this.configuration = configuration;
            if (string.IsNullOrEmpty(configuration))
            {
                this.validServers.Clear();
                return;
            }

            try
            {
                var xml = XDocument.Parse(configuration);
                XElement validServersXml = xml.Root.Element("ValidServers");
                if(validServersXml == null)
                {
                    this.validServers.Clear();
                    return;
                }
                foreach (var validServer in validServersXml.Elements("Server"))
                {
                    if (!string.IsNullOrEmpty(validServer.Value))
                        this.validServers.Add(validServer.Value.ToLower());
                }
            }
            catch (Exception ex)
            {
                string message = "There is an error in the extension configuration. Find rsreportserver.config file and check if the Configuration element under 'JC SharePoint Delivery' node contains valid XML.";
                throw new Exception(message, ex);
            }
        }


        public Setting[] ValidateUserData(Setting[] settings)
        {
            SubscriptionData sd = new SharePointDelivery.SubscriptionData();
            sd.Server = settings[0].Value;
            sd.path = settings[1].Value;
            sd.file = settings[2].Value;
            sd.renderingFormat = settings[3].Value;

            if(this.validServers.Count == 0)
            {
                if(sd.Server.ToLower().StartsWith("http://") && sd.Server.ToLower().StartsWith("https://"))
                {
                    settings[0].Error = "Invalid server name. The name must start with https:// or http://";
                }
            }
            else // preconfigured values
            {
                if (!validServers.Contains(sd.Server.ToLower()))
                    settings[0].Error = "Invalid server name. It must be one of values preconfigured in rsreportserver.config file.";
            }            

            return settings; // we are not checking anything right now
        }      
    }
}
