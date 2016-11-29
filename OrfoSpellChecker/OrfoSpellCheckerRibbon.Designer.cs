namespace OrfoSpellChecker
{
    partial class OrfoSpellCheckerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public OrfoSpellCheckerRibbon()
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
            this.OrfoSpellCheckerTab = this.Factory.CreateRibbonTab();
            this.OrfoSpellCheckerTabGroup = this.Factory.CreateRibbonGroup();
            this.OSCCheckAll = this.Factory.CreateRibbonButton();
            this.OSCCheckCurrentTab = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.OSCAutoCheck = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.OrfoSpellCheckerTab.SuspendLayout();
            this.OrfoSpellCheckerTabGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // OrfoSpellCheckerTab
            // 
            this.OrfoSpellCheckerTab.Groups.Add(this.OrfoSpellCheckerTabGroup);
            this.OrfoSpellCheckerTab.Label = "OrfoSpellChecker";
            this.OrfoSpellCheckerTab.Name = "OrfoSpellCheckerTab";
            // 
            // OrfoSpellCheckerTabGroup
            // 
            this.OrfoSpellCheckerTabGroup.Items.Add(this.OSCCheckAll);
            this.OrfoSpellCheckerTabGroup.Items.Add(this.OSCCheckCurrentTab);
            this.OrfoSpellCheckerTabGroup.Items.Add(this.separator1);
            this.OrfoSpellCheckerTabGroup.Items.Add(this.OSCAutoCheck);
            this.OrfoSpellCheckerTabGroup.Label = "OrfoSpellCheck";
            this.OrfoSpellCheckerTabGroup.Name = "OrfoSpellCheckerTabGroup";
            // 
            // OSCCheckAll
            // 
            this.OSCCheckAll.Label = "Check All Book";
            this.OSCCheckAll.Name = "OSCCheckAll";
            this.OSCCheckAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OSCCheckAll_Click);
            // 
            // OSCCheckCurrentTab
            // 
            this.OSCCheckCurrentTab.Label = "Check Current Tab";
            this.OSCCheckCurrentTab.Name = "OSCCheckCurrentTab";
            this.OSCCheckCurrentTab.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OSCCheckCurrentTab_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // OSCAutoCheck
            // 
            this.OSCAutoCheck.Label = "Auto Checking";
            this.OSCAutoCheck.Name = "OSCAutoCheck";
            this.OSCAutoCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OSCAutoCheck_Click);
            // 
            // OrfoSpellCheckerRibbon
            // 
            this.Name = "OrfoSpellCheckerRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.OrfoSpellCheckerTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OrfoSpellCheckerRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.OrfoSpellCheckerTab.ResumeLayout(false);
            this.OrfoSpellCheckerTab.PerformLayout();
            this.OrfoSpellCheckerTabGroup.ResumeLayout(false);
            this.OrfoSpellCheckerTabGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab OrfoSpellCheckerTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup OrfoSpellCheckerTabGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OSCCheckAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OSCCheckCurrentTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton OSCAutoCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal OrfoSpellCheckerRibbon OrfoSpellCheckerRibbon
        {
            get { return this.GetRibbon<OrfoSpellCheckerRibbon>(); }
        }
    }
}
