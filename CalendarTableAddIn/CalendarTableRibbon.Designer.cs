namespace CalendarTableAddIn
{
    partial class CalendarTableRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CalendarTableRibbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.TabInsertCalendarTable = this.Factory.CreateRibbonTab();
            this.GroupInsertCalendarTables = this.Factory.CreateRibbonGroup();
            this.ButtonInsertCalendarTable = this.Factory.CreateRibbonButton();
            this.TabInsertCalendarTable.SuspendLayout();
            this.GroupInsertCalendarTables.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabInsertCalendarTable
            // 
            this.TabInsertCalendarTable.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabInsertCalendarTable.ControlId.OfficeId = "TabInsert";
            this.TabInsertCalendarTable.Groups.Add(this.GroupInsertCalendarTables);
            this.TabInsertCalendarTable.Label = "TabInsert";
            this.TabInsertCalendarTable.Name = "TabInsertCalendarTable";
            // 
            // GroupInsertCalendarTables
            // 
            ribbonDialogLauncherImpl1.ScreenTip = "Select Month";
            this.GroupInsertCalendarTables.DialogLauncher = ribbonDialogLauncherImpl1;
            this.GroupInsertCalendarTables.Items.Add(this.ButtonInsertCalendarTable);
            this.GroupInsertCalendarTables.Label = "Calendar Table";
            this.GroupInsertCalendarTables.Name = "GroupInsertCalendarTables";
            this.GroupInsertCalendarTables.Position = this.Factory.RibbonPosition.AfterOfficeId("GroupInsertTables");
            this.GroupInsertCalendarTables.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GroupInsertCalendarTables_DialogLauncherClick);
            // 
            // ButtonInsertCalendarTable
            // 
            this.ButtonInsertCalendarTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonInsertCalendarTable.Image = global::CalendarTableAddIn.Properties.Resources.icon;
            this.ButtonInsertCalendarTable.Label = "Current Month";
            this.ButtonInsertCalendarTable.Name = "ButtonInsertCalendarTable";
            this.ButtonInsertCalendarTable.ShowImage = true;
            this.ButtonInsertCalendarTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.TabInsertCalendarTable);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.TabInsertCalendarTable.ResumeLayout(false);
            this.TabInsertCalendarTable.PerformLayout();
            this.GroupInsertCalendarTables.ResumeLayout(false);
            this.GroupInsertCalendarTables.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabInsertCalendarTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupInsertCalendarTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonInsertCalendarTable;
    }

    partial class ThisRibbonCollection
    {
        internal CalendarTableRibbon Ribbon
        {
            get { return this.GetRibbon<CalendarTableRibbon>(); }
        }
    }
}
