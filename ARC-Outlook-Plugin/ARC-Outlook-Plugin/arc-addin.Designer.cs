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
            this.arc_tab = this.Factory.CreateRibbonTab();
            this.arc_btn_group = this.Factory.CreateRibbonGroup();
            this.btn_arc = this.Factory.CreateRibbonButton();
            this.arc_tab.SuspendLayout();
            this.arc_btn_group.SuspendLayout();
            this.SuspendLayout();
            // 
            // arc_tab
            // 
            this.arc_tab.Groups.Add(this.arc_btn_group);
            resources.ApplyResources(this.arc_tab, "arc_tab");
            this.arc_tab.Name = "arc_tab";
            // 
            // arc_btn_group
            // 
            this.arc_btn_group.Items.Add(this.btn_arc);
            resources.ApplyResources(this.arc_btn_group, "arc_btn_group");
            this.arc_btn_group.Name = "arc_btn_group";
            // 
            // btn_arc
            // 
            this.btn_arc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.btn_arc, "btn_arc");
            this.btn_arc.Name = "btn_arc";
            this.btn_arc.ShowImage = true;
            this.btn_arc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_arc_Click);
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
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup arc_btn_group;
        private Microsoft.Office.Tools.Ribbon.RibbonTab arc_tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_arc;
    }

    partial class ThisRibbonCollection
    {
        internal arc_addin arc_addin
        {
            get { return this.GetRibbon<arc_addin>(); }
        }
    }
}
