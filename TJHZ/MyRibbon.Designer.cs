namespace TJHZ
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpEmails = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnCreateEmail = this.Factory.CreateRibbonButton();
            this.btnAddAttachment = this.Factory.CreateRibbonButton();
            this.btnGetEmails = this.Factory.CreateRibbonButton();
            this.btnSaveAttachments = this.Factory.CreateRibbonButton();
            this.grpContacts = this.Factory.CreateRibbonGroup();
            this.btnFindContact = this.Factory.CreateRibbonButton();
            this.btnAddContact = this.Factory.CreateRibbonButton();
            this.btnDeleteContact = this.Factory.CreateRibbonButton();
            this.grpCalendar = this.Factory.CreateRibbonGroup();
            this.btnCreateAppointment = this.Factory.CreateRibbonButton();
            this.btnDeleteAppointment = this.Factory.CreateRibbonButton();
            this.btnCreateCalendar = this.Factory.CreateRibbonButton();
            this.grpFolders = this.Factory.CreateRibbonGroup();
            this.btnCreateFolder = this.Factory.CreateRibbonButton();
            this.btnSelectFolder = this.Factory.CreateRibbonButton();
            this.btnDeleteFolder = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpEmails.SuspendLayout();
            this.grpContacts.SuspendLayout();
            this.grpCalendar.SuspendLayout();
            this.grpFolders.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpEmails);
            this.tab1.Groups.Add(this.grpContacts);
            this.tab1.Groups.Add(this.grpCalendar);
            this.tab1.Groups.Add(this.grpFolders);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpEmails
            // 
            this.grpEmails.Items.Add(this.button1);
            this.grpEmails.Items.Add(this.btnCreateEmail);
            this.grpEmails.Items.Add(this.btnAddAttachment);
            this.grpEmails.Items.Add(this.btnGetEmails);
            this.grpEmails.Items.Add(this.btnSaveAttachments);
            this.grpEmails.Label = "Emails";
            this.grpEmails.Name = "grpEmails";
            // 
            // button1
            // 
            this.button1.Label = "Get Selected Text";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetSelectedText_Click);
            // 
            // btnCreateEmail
            // 
            this.btnCreateEmail.Label = "Create Email";
            this.btnCreateEmail.Name = "btnCreateEmail";
            this.btnCreateEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateEmail_Click);
            // 
            // btnAddAttachment
            // 
            this.btnAddAttachment.Label = "Send Email with Attachment";
            this.btnAddAttachment.Name = "btnAddAttachment";
            this.btnAddAttachment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddAttachment_Click);
            // 
            // btnGetEmails
            // 
            this.btnGetEmails.Label = "Get New Emails";
            this.btnGetEmails.Name = "btnGetEmails";
            this.btnGetEmails.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGetEmails_Click);
            // 
            // btnSaveAttachments
            // 
            this.btnSaveAttachments.Label = "Save Attachments";
            this.btnSaveAttachments.Name = "btnSaveAttachments";
            this.btnSaveAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveAttachments_Click);
            // 
            // grpContacts
            // 
            this.grpContacts.Items.Add(this.btnFindContact);
            this.grpContacts.Items.Add(this.btnAddContact);
            this.grpContacts.Items.Add(this.btnDeleteContact);
            this.grpContacts.Label = "Contacts";
            this.grpContacts.Name = "grpContacts";
            // 
            // btnFindContact
            // 
            this.btnFindContact.Label = "Find Contact";
            this.btnFindContact.Name = "btnFindContact";
            this.btnFindContact.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFindContact_Click);
            // 
            // btnAddContact
            // 
            this.btnAddContact.Label = "Add Contact";
            this.btnAddContact.Name = "btnAddContact";
            this.btnAddContact.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddContact_Click);
            // 
            // btnDeleteContact
            // 
            this.btnDeleteContact.Label = "Delete Contact";
            this.btnDeleteContact.Name = "btnDeleteContact";
            this.btnDeleteContact.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteContact_Click);
            // 
            // grpCalendar
            // 
            this.grpCalendar.Items.Add(this.btnCreateAppointment);
            this.grpCalendar.Items.Add(this.btnDeleteAppointment);
            this.grpCalendar.Items.Add(this.btnCreateCalendar);
            this.grpCalendar.Label = "Calendar";
            this.grpCalendar.Name = "grpCalendar";
            // 
            // btnCreateAppointment
            // 
            this.btnCreateAppointment.Label = "Create Appointment";
            this.btnCreateAppointment.Name = "btnCreateAppointment";
            this.btnCreateAppointment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateAppointment_Click);
            // 
            // btnDeleteAppointment
            // 
            this.btnDeleteAppointment.Label = "Delete Appointment";
            this.btnDeleteAppointment.Name = "btnDeleteAppointment";
            this.btnDeleteAppointment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteAppointment_Click);
            // 
            // btnCreateCalendar
            // 
            this.btnCreateCalendar.Label = "Create Calendar";
            this.btnCreateCalendar.Name = "btnCreateCalendar";
            this.btnCreateCalendar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateCalendar_Click);
            // 
            // grpFolders
            // 
            this.grpFolders.Items.Add(this.btnCreateFolder);
            this.grpFolders.Items.Add(this.btnSelectFolder);
            this.grpFolders.Items.Add(this.btnDeleteFolder);
            this.grpFolders.Label = "Folders";
            this.grpFolders.Name = "grpFolders";
            // 
            // btnCreateFolder
            // 
            this.btnCreateFolder.Label = "Create Folder";
            this.btnCreateFolder.Name = "btnCreateFolder";
            this.btnCreateFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateFolder_Click);
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Label = "Select Folder";
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectFolder_Click);
            // 
            // btnDeleteFolder
            // 
            this.btnDeleteFolder.Label = "Delete Folder";
            this.btnDeleteFolder.Name = "btnDeleteFolder";
            this.btnDeleteFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteFolder_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpEmails.ResumeLayout(false);
            this.grpEmails.PerformLayout();
            this.grpContacts.ResumeLayout(false);
            this.grpContacts.PerformLayout();
            this.grpCalendar.ResumeLayout(false);
            this.grpCalendar.PerformLayout();
            this.grpFolders.ResumeLayout(false);
            this.grpFolders.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEmails;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateEmail;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddAttachment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGetEmails;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveAttachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpContacts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFindContact;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddContact;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteContact;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpCalendar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateAppointment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteAppointment;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateCalendar;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFolders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteFolder;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
