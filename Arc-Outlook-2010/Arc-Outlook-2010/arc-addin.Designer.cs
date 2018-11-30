namespace Arc_Outlook_2010
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
            this.arc_tab = this.Factory.CreateRibbonTab();
            this.arc_btn_group = this.Factory.CreateRibbonGroup();
            this.btn_arc = this.Factory.CreateRibbonButton();
            this.syncBtn = this.Factory.CreateRibbonButton();
            this.group_setting = this.Factory.CreateRibbonGroup();
            this.cleanDataBtn = this.Factory.CreateRibbonButton();
            this.inforBtn = this.Factory.CreateRibbonButton();
            this.arcWebBtn = this.Factory.CreateRibbonButton();
            this.arc_tab.SuspendLayout();
            this.arc_btn_group.SuspendLayout();
            this.group_setting.SuspendLayout();
            this.SuspendLayout();
            // 
            // arc_tab
            // 
            this.arc_tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.arc_tab.ControlId.OfficeId = "arc_tab";
            this.arc_tab.Groups.Add(this.arc_btn_group);
            this.arc_tab.Groups.Add(this.group_setting);
            resources.ApplyResources(this.arc_tab, "arc_tab");
            this.arc_tab.Name = "arc_tab";
            // 
            // arc_btn_group
            // 
            this.arc_btn_group.Items.Add(this.btn_arc);
            this.arc_btn_group.Items.Add(this.syncBtn);
            resources.ApplyResources(this.arc_btn_group, "arc_btn_group");
            this.arc_btn_group.Name = "arc_btn_group";
            // 
            // btn_arc
            // 
            this.btn_arc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.btn_arc, "btn_arc");
            this.btn_arc.Name = "btn_arc";
            this.btn_arc.ShowImage = true;
            // 
            // syncBtn
            // 
            this.syncBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.syncBtn, "syncBtn");
            this.syncBtn.Name = "syncBtn";
            this.syncBtn.ShowImage = true;
            // 
            // group_setting
            // 
            this.group_setting.Items.Add(this.cleanDataBtn);
            this.group_setting.Items.Add(this.inforBtn);
            this.group_setting.Items.Add(this.arcWebBtn);
            resources.ApplyResources(this.group_setting, "group_setting");
            this.group_setting.Name = "group_setting";
            // 
            // cleanDataBtn
            // 
            resources.ApplyResources(this.cleanDataBtn, "cleanDataBtn");
            this.cleanDataBtn.Name = "cleanDataBtn";
            this.cleanDataBtn.ShowImage = true;
            // 
            // inforBtn
            // 
            resources.ApplyResources(this.inforBtn, "inforBtn");
            this.inforBtn.Name = "inforBtn";
            this.inforBtn.ShowImage = true;
            // 
            // arcWebBtn
            // 
            this.arcWebBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.arcWebBtn, "arcWebBtn");
            this.arcWebBtn.Name = "arcWebBtn";
            this.arcWebBtn.ShowImage = true;
            // 
            // arc_addin
            // 
            this.Name = "arc_addin";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.arc_tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.arc_addin_Load);
            this.arc_tab.ResumeLayout(false);
            this.arc_tab.PerformLayout();
            this.arc_btn_group.ResumeLayout(false);
            this.arc_btn_group.PerformLayout();
            this.group_setting.ResumeLayout(false);
            this.group_setting.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup arc_btn_group;
        private Microsoft.Office.Tools.Ribbon.RibbonTab arc_tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_arc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton syncBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_setting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton cleanDataBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton inforBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton arcWebBtn;
    }

    partial class ThisRibbonCollection
    {
        internal arc_addin arc_addin
        {
            get { return this.GetRibbon<arc_addin>(); }
        }
    }
}
