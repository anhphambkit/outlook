namespace ARC_Outlook_Plugin
{
    partial class arc_addin : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public arc_addin()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(arc_addin));
            this.tab2 = this.Factory.CreateRibbonTab();
            this.arc_btn_group = this.Factory.CreateRibbonGroup();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tab3 = this.Factory.CreateRibbonTab();
            this.tab2.SuspendLayout();
            this.tab1.SuspendLayout();
            this.tab3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab2.ControlId.OfficeId = "tab2";
            this.tab2.Groups.Add(this.arc_btn_group);
            resources.ApplyResources(this.tab2, "tab2");
            this.tab2.Name = "tab2";
            // 
            // arc_btn_group
            // 
            resources.ApplyResources(this.arc_btn_group, "arc_btn_group");
            this.arc_btn_group.Name = "arc_btn_group";
            // 
            // tab1
            // 
            resources.ApplyResources(this.tab1, "tab1");
            this.tab1.Name = "tab1";
            // 
            // tab3
            // 
            resources.ApplyResources(this.tab3, "tab3");
            this.tab3.Name = "tab3";
            // 
            // arc_addin
            // 
            this.Name = "arc_addin";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Tabs.Add(this.tab3);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.arc_addin_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab3.ResumeLayout(false);
            this.tab3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup arc_btn_group;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab3;
    }

    partial class ThisRibbonCollection
    {
        internal arc_addin arc_addin
        {
            get { return this.GetRibbon<arc_addin>(); }
        }
    }
}
