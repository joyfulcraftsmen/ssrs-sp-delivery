using System;
using System.IO;
using System.Web;
using System.Text;
using System.Web.UI;
using System.Collections;
using System.Globalization;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using Microsoft.ReportingServices.Interfaces;
using System.Xml;

namespace JC.SSRS.Extensions.SharePointDelivery
{
    // ServerDeliveryUIProvider implements a UserControl that plugs into the Reporting Services Report Manager application.
    public class SharePointDeliveryExtensionUI : System.Web.UI.WebControls.WebControl, ISubscriptionBaseUIUserControl
    {
        // Variables used to store information about the subscription being created
        private string m_path = "";
        private string m_file = "";
        private string m_server = "";
        private string m_renderingFormat = "PDF";

        // Labels for UI controls
        internal const string SERVERLABEL = "Select a SharePoint server: ";
        internal const string PATHLABEL = "Path where you want to export: ";
        internal const string FILELABEL = "Name of the exported file: ";
        internal const string RENDERINGFORMATLABEL = "Rendering format: ";

        // IDs used to refer to controls on the UI page.
        internal const string SERVERCONTROLID = "SERVERTEXTBOX";
        internal const string PATHCONTROLID = "PATHTEXTBOX";
        internal const string FILECONTROLID = "FILETEXTBOX";
        internal const string RENDERINGFORMATCONTROLID = "RENDERINGFORMATSELECTBOX";

        // Strings to enable validation to occur (client side validation needs JavaScript)
        internal const string LANGUAGEATTRIBUTE = "language";
        internal const string SCRIPTDEFAULTLANGUAGE = "Javascript";
        internal const string SCRIPTTYPEATTRIBUTE = "type";
        internal const string SCRIPTDEFAULTTYPE = "text/Javascript";
        internal const string SCRIPTTAG = "script";
        //internal const string ONCLICK = "onclick";

        // Used to keep track of whether we have values specified by the user
        private bool m_hasUserData;

        #region Controls
        // Controls used on the UI page
        private LiteralControl m_validatorScript = new LiteralControl();

        // HTML table variables
        private HtmlTable m_outerTable = new HtmlTable();
        private HtmlTableRow m_currentRow;
        private HtmlTableCell m_currentCell;
        private HtmlGenericControl m_pageLevelScript = new HtmlGenericControl(SCRIPTTAG);

        // Labels for controls
        private Label m_serverLabel = new Label();
        private Label m_pathLabel = new Label();
        private Label m_fileLabel = new Label();
        private Label m_renderingFormatLabel = new Label();

        // Control types
        private TextBox m_serverTextBox = new TextBox();
        private TextBox m_pathTextBox = new TextBox();
        private TextBox m_fileTextBox = new TextBox();
        private DropDownList m_renderingFormatSelectBox = new DropDownList();

        // Placeholders for validators
        private PlaceHolder m_invalidServerName = new PlaceHolder();
        private PlaceHolder m_invalidPath = new PlaceHolder();
        private PlaceHolder m_invalidFile = new PlaceHolder();
        private PlaceHolder m_invalidRenderingFormat = new PlaceHolder();

        // Field validators for UI
        private ServerNameValidator m_serverRequired = new ServerNameValidator();
        private RequiredFieldValidator m_pathRequired = new RequiredFieldValidator();
        private RequiredFieldValidator m_fileRequired = new RequiredFieldValidator();
        private RequiredFieldValidator m_renderingFormatRequired = new RequiredFieldValidator();

        #endregion

        // Provider constructor
        public SharePointDeliveryExtensionUI()
        {

            this.Init += new EventHandler(Control_Init);
            this.Load += new EventHandler(Control_Load);
            this.PreRender += new EventHandler(Control_PreRender);
        }

        #region Event Handlers

        private void Control_PreRender(object sender, EventArgs args)
        {
            // Use this event to enable/disable controls based on the 
            // user's selection before the control is rendered.
            // The ServerDeliverySample does not use this event handler.
        }

        /// <summary>Perform all step needed when the control has been loaded</summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void Control_Load(object sender, EventArgs args)
        {
            if (!Page.IsPostBack)
            {
                //if you have non-required Extension settings, initialize the values of the controls here
            }
        }

        /// <summary>
        /// Initialize the control
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void Control_Init(object sender, EventArgs args)
        {
            Controls.Add(m_pageLevelScript);

            // Create page level script
            m_pageLevelScript.Attributes.Add(LANGUAGEATTRIBUTE, SCRIPTDEFAULTLANGUAGE);
            m_pageLevelScript.Attributes.Add(SCRIPTTYPEATTRIBUTE, SCRIPTDEFAULTTYPE);
            Controls.Add(m_validatorScript);
            SetValidatorScript();

            // Build a table row for selecting a server from the serversDropDownList
            #region Server DropDown Row
            m_currentRow = new HtmlTableRow();
            m_outerTable.Rows.Add(m_currentRow);

            // Create the first cell that contains the label for the server drop down list
            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "10%";

            //TODO: Specify label here
            this.m_serverLabel.Text = HttpUtility.HtmlEncode(SERVERLABEL);
            m_currentCell.Controls.Add(this.m_serverLabel);

            // Add the cell that contains the drop down list of server names
            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "25%";
            //TODO: Add control here
            // this.m_serverTextBox.SelectedIndex = 0;
            //TODO: Include list of servers here...
            //TODO: Define default values here (merge UserData and ServerInfo)

            this.m_serverTextBox.ID = SERVERCONTROLID;
            this.m_serverTextBox.Style.Add("width", "400px");
            this.m_serverTextBox.Style.Add("height", "20px");
            this.m_serverTextBox.Style.Add("padding", "5px");
            this.m_serverTextBox.Style.Add("margin-left", "20px");
            this.m_serverTextBox.Style.Add("margin-bottom", "5px");
            this.m_serverTextBox.Style.Add("font-family", "\"Segoe UI\", \"Helvetica Neue\", Helvetica, Arial, sans-serif");
            this.m_serverTextBox.Style.Add("font-size", "14px");
            m_currentCell.Controls.Add(this.m_serverTextBox);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "40%";

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "100%";

            m_currentCell.Controls.Add(this.m_invalidServerName);
            this.m_serverRequired.Display = ValidatorDisplay.Dynamic;
            this.m_serverRequired.ControlToValidate = SERVERCONTROLID;
            String serverRequiredError = "SharePoint server URL is required.";
            this.m_serverRequired.Controls.Add(ErrorMessage(serverRequiredError, true));
            this.m_invalidServerName.Controls.Add(this.m_serverRequired);
            #endregion

            // Build a table row for entering a path
            #region Page Width Textbox Row
            m_currentRow = new HtmlTableRow();
            m_outerTable.Rows.Add(m_currentRow);

            // Add cell for page width label
            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "10%";

            // Add label
            this.m_pathLabel.Text = HttpUtility.HtmlEncode(PATHLABEL);
            m_currentCell.Controls.Add(this.m_pathLabel);

            // Add text box for entering value
            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "25%";

            this.m_pathTextBox.Text = System.Convert.ToString(
                this.m_path,
                System.Globalization.CultureInfo.InvariantCulture);

            m_pathTextBox.ID = PATHCONTROLID;
            m_pathTextBox.Style.Add("font-family", "Verdana, Sans-Serif");
            m_pathTextBox.Style.Add("font-size", "x-small");

            this.m_pathTextBox.ID = PATHCONTROLID;
            this.m_pathTextBox.Style.Add("width", "400px");
            this.m_pathTextBox.Style.Add("height", "20px");
            this.m_pathTextBox.Style.Add("padding", "5px");
            this.m_pathTextBox.Style.Add("margin-left", "20px");
            this.m_pathTextBox.Style.Add("margin-bottom", "5px");
            m_currentCell.Controls.Add(this.m_pathTextBox);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "40%";

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "100%";

            // Add validator here
            m_currentCell.Controls.Add(this.m_invalidPath);
            this.m_pathRequired.Display = ValidatorDisplay.Dynamic;
            this.m_pathRequired.ControlToValidate = PATHCONTROLID;
            String pathRequiredError = "Path is required.";
            this.m_pathRequired.Controls.Add(ErrorMessage(pathRequiredError, true));
            this.m_invalidPath.Controls.Add(this.m_pathRequired);

            RequiredFieldValidator vgtzValidator1 = new RequiredFieldValidator();
            vgtzValidator1.Display = ValidatorDisplay.Dynamic;
            vgtzValidator1.ControlToValidate = PATHCONTROLID;
            vgtzValidator1.Controls.Add(ErrorMessage("This value is required.", true));
            this.m_invalidPath.Controls.Add(vgtzValidator1);

            #endregion

            // Build a table row for Entering a file name
            #region Page Height Textbox Row
            m_currentRow = new HtmlTableRow();
            m_outerTable.Rows.Add(m_currentRow);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "10%";

            this.m_fileLabel.Text = HttpUtility.HtmlEncode(FILELABEL);
            m_currentCell.Controls.Add(this.m_fileLabel);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "25%";

            this.m_fileTextBox.Text = System.Convert.ToString(this.m_file,
                System.Globalization.CultureInfo.InvariantCulture);

            m_fileTextBox.ID = FILECONTROLID;
            m_fileTextBox.Style.Add("font-family", "Verdana, Sans-Serif");
            m_fileTextBox.Style.Add("font-size", "x-small");

            this.m_fileTextBox.ID = FILECONTROLID;
            this.m_fileTextBox.Style.Add("width", "400px");
            this.m_fileTextBox.Style.Add("height", "20px");
            this.m_fileTextBox.Style.Add("padding", "5px");
            this.m_fileTextBox.Style.Add("margin-left", "20px");
            this.m_fileTextBox.Style.Add("margin-bottom", "5px");
            m_currentCell.Controls.Add(this.m_fileTextBox);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "40%";

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "100%";

            m_currentCell.Controls.Add(this.m_invalidFile);
            this.m_fileRequired.Display = ValidatorDisplay.Dynamic;
            this.m_fileRequired.ControlToValidate = FILECONTROLID;
            String fileRequired = "File name is required.";
            this.m_fileRequired.Controls.Add(ErrorMessage(fileRequired, true));
            this.m_invalidFile.Controls.Add(this.m_fileRequired);

            RequiredFieldValidator vgtzValidator2 = new RequiredFieldValidator();
            vgtzValidator2.Display = ValidatorDisplay.Dynamic;
            vgtzValidator2.ControlToValidate = FILECONTROLID;
            vgtzValidator2.Controls.Add(ErrorMessage("This value is required.", true));
            this.m_invalidFile.Controls.Add(vgtzValidator2);

            #endregion

            // Build a table row for Entering a rendering format
            #region Rendering Format Textbox Row
            m_currentRow = new HtmlTableRow();
            m_outerTable.Rows.Add(m_currentRow);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "10%";

            foreach (var renderingExtension in this.ReportServerInformation.RenderingExtension)
            {
                if (renderingExtension.Visible)
                {
                    ListItem li = new ListItem(renderingExtension.LocalizedName, renderingExtension.Name);
                    m_renderingFormatSelectBox.Items.Add(li);
                }
            }
            this.m_renderingFormatSelectBox.SelectedIndex = 0;
            this.m_renderingFormatLabel.Text = HttpUtility.HtmlEncode(RENDERINGFORMATLABEL);
            m_currentCell.Controls.Add(this.m_renderingFormatLabel);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "25%";

            this.m_renderingFormatSelectBox.Text = System.Convert.ToString(this.m_renderingFormat,
                System.Globalization.CultureInfo.InvariantCulture);

            m_renderingFormatSelectBox.ID = RENDERINGFORMATCONTROLID;
            m_renderingFormatSelectBox.Style.Add("font-family", "Verdana, Sans-Serif");
            m_renderingFormatSelectBox.Style.Add("font-size", "x-small");
            this.m_renderingFormatSelectBox.Style.Add("width", "400px");
            // this.m_renderingFormatSelectBox.Style.Add("height", "20px");
            this.m_renderingFormatSelectBox.Style.Add("padding", "5px");
            this.m_renderingFormatSelectBox.Style.Add("margin-left", "20px");
            this.m_renderingFormatSelectBox.Style.Add("margin-bottom", "5px");

            this.m_renderingFormatSelectBox.ID = RENDERINGFORMATCONTROLID;
            m_currentCell.Controls.Add(this.m_renderingFormatSelectBox);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "40%";

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "100%";

            m_currentCell.Controls.Add(this.m_invalidRenderingFormat);
            this.m_renderingFormatRequired.Display = ValidatorDisplay.Dynamic;
            this.m_renderingFormatRequired.ControlToValidate = RENDERINGFORMATCONTROLID;
            String renderingFormatRequired = "Specify a value for the page height.";
            this.m_renderingFormatRequired.Controls.Add(ErrorMessage(renderingFormatRequired, true));
            this.m_invalidRenderingFormat.Controls.Add(this.m_renderingFormatRequired);

            RequiredFieldValidator vgtzValidator3 = new RequiredFieldValidator();
            vgtzValidator3.Display = ValidatorDisplay.Dynamic;
            vgtzValidator3.ControlToValidate = RENDERINGFORMATCONTROLID;
            vgtzValidator3.Controls.Add(ErrorMessage("This value is required.", true));
            this.m_invalidRenderingFormat.Controls.Add(vgtzValidator3);

            #endregion

            m_outerTable.Attributes.Add("class", "msrs-normal");
            m_outerTable.CellPadding = 0;
            m_outerTable.CellSpacing = 0;
            m_outerTable.Width = "100%";
            Controls.Add(m_outerTable);

        }

        #region VALIDATORSCRIPTFUNCTION
        private const string VALIDATORSCRIPTFUNCTION =
           @"<script language='Javascript' type='text/Javascript'>
                function validateServerName(source, args)
                {
                  args.IsValid = true;
                }
                </script>
                ";
        #endregion
        private void SetValidatorScript()
        {
            m_validatorScript.Text = VALIDATORSCRIPTFUNCTION;
        }

        protected Control ErrorMessage(string error, bool noWrap)
        {
            string imgUrl = Page.Request.ApplicationPath + "/images/line_err1.gif";
            HtmlImage htmlImg = new HtmlImage();
            htmlImg.Src = imgUrl;
            htmlImg.Alt = "Value specified contains an error.";

            HtmlTable tbl = new HtmlTable();
            tbl.Rows.Add(new HtmlTableRow());
            HtmlTableCell cell = new HtmlTableCell();
            cell.VAlign = "middle";
            //Note: you can reuse the Report Manger style sheet.
            cell.Attributes.Add("class", "msrs-validationerror");
            cell.Controls.Add(htmlImg);
            tbl.Rows[0].Cells.Add(cell);
            cell = new HtmlTableCell();
            cell.NoWrap = noWrap;
            cell.VAlign = "middle";
            cell.Attributes.Add("class", "msrs-validationerror");
            cell.Controls.Add(new LiteralControl(HttpUtility.HtmlEncode(error)));
            tbl.Rows[0].Cells.Add(cell);
            return tbl;
        }

        #endregion

        #region IExtension methods
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public String LocalizedName
        {
            get
            {
                return "JC SharePoint Delivery Extension";
            }
        }

        private string m_configuration { get; set; } = null;

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public void SetConfiguration(String configuration)
        {
            CultureInfo info = System.Threading.Thread.CurrentThread.CurrentCulture;
            try
            {
                this.m_configuration = configuration;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, null);
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = info;
            }
        }

        #endregion

        #region ISubscriptionBaseUIUserControl methods

        private bool m_isPrivilegedUser;

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public bool IsPrivilegedUser
        {
            get
            {
                return m_isPrivilegedUser;
            }
            set
            {
                m_isPrivilegedUser = value;
            }
        }

        // Validate that all selected information is correct
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public bool Validate()
        {
            // Nothing additional to validate
            return true;
        }

        // Get and Set the user data
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public Setting[] UserData
        {
            get
            {
                SubscriptionData data = new SubscriptionData();
                data.Server = this.m_serverTextBox.Text;
                data.path = this.m_pathTextBox.Text;
                data.file = this.m_fileTextBox.Text;
                data.renderingFormat = this.m_renderingFormatSelectBox.SelectedItem.Value;

                return data.ToSettingArray();
            }

            set
            {
                this.m_hasUserData = true;

                SubscriptionData data = new SubscriptionData();
                data.FromSettings(value);

                this.m_server = data.Server;
                this.m_serverTextBox.Text = this.m_server;

                this.m_path = data.path;
                this.m_pathTextBox.Text = this.m_path;

                this.m_file = data.file;
                this.m_fileTextBox.Text = this.m_file;

                this.m_renderingFormat = data.renderingFormat;
                this.m_renderingFormatSelectBox.SelectedValue = this.m_renderingFormat;

                bool found = false;
                Setting[] serverSettings = m_rsInformation.ServerSettings;
                Setting servers = null;
                foreach (Setting s in serverSettings)
                {
                    if (s.Name.Equals(SubscriptionData.SERVER))
                    {
                        servers = s;
                    }
                }
            }
        }

        // Get the description that displays for the subscription
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public String Description
        {
            get
            {
                return "Print report to " + this.m_serverTextBox.Text + ".";
            }
        }

        private IDeliveryReportServerInformation m_rsInformation;
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public IDeliveryReportServerInformation ReportServerInformation
        {
            set
            {
                m_rsInformation = value;
            }
            get
            {
                return m_rsInformation;
            }
        }

        #endregion

    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1501:AvoidExcessiveInheritance")]
    internal class ServerNameValidator : CustomValidator
    {
        public ServerNameValidator()
            : base()
        {
            ServerValidate += new ServerValidateEventHandler(Validate_Server);

        }

        private void Validate_Server(object source, ServerValidateEventArgs args)
        {
            TextBox tb = (TextBox)FindControl(((CustomValidator)source).ControlToValidate);
            if (tb.Text.ToLower().StartsWith("https://") || tb.Text.ToLower().StartsWith("http://"))
                args.IsValid = true;
            else
                args.IsValid = false;
        }
    }
}



