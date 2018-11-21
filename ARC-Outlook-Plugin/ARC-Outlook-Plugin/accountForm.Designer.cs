namespace ARC_Outlook_Plugin
{
    partial class accountForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(accountForm));
            this.formGeneralPanel = new System.Windows.Forms.Panel();
            this.leftLogoPanel = new System.Windows.Forms.Panel();
            this.rightFormPanel = new System.Windows.Forms.Panel();
            this.logoPlugin = new System.Windows.Forms.PictureBox();
            this.formGeneralPanel.SuspendLayout();
            this.leftLogoPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.logoPlugin)).BeginInit();
            this.SuspendLayout();
            // 
            // formGeneralPanel
            // 
            this.formGeneralPanel.AutoSize = true;
            this.formGeneralPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.formGeneralPanel.BackColor = System.Drawing.Color.Transparent;
            this.formGeneralPanel.Controls.Add(this.rightFormPanel);
            this.formGeneralPanel.Controls.Add(this.leftLogoPanel);
            this.formGeneralPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.formGeneralPanel.Location = new System.Drawing.Point(4, 4);
            this.formGeneralPanel.Margin = new System.Windows.Forms.Padding(33, 32, 33, 32);
            this.formGeneralPanel.Name = "formGeneralPanel";
            this.formGeneralPanel.Padding = new System.Windows.Forms.Padding(100, 86, 100, 86);
            this.formGeneralPanel.Size = new System.Drawing.Size(1229, 593);
            this.formGeneralPanel.TabIndex = 11;
            // 
            // leftLogoPanel
            // 
            this.leftLogoPanel.Controls.Add(this.logoPlugin);
            this.leftLogoPanel.Dock = System.Windows.Forms.DockStyle.Left;
            this.leftLogoPanel.Location = new System.Drawing.Point(100, 86);
            this.leftLogoPanel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.leftLogoPanel.Name = "leftLogoPanel";
            this.leftLogoPanel.Size = new System.Drawing.Size(492, 421);
            this.leftLogoPanel.TabIndex = 0;
            // 
            // rightFormPanel
            // 
            this.rightFormPanel.BackColor = System.Drawing.Color.White;
            this.rightFormPanel.Dock = System.Windows.Forms.DockStyle.Right;
            this.rightFormPanel.Location = new System.Drawing.Point(653, 86);
            this.rightFormPanel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.rightFormPanel.Name = "rightFormPanel";
            this.rightFormPanel.Padding = new System.Windows.Forms.Padding(13, 0, 13, 12);
            this.rightFormPanel.Size = new System.Drawing.Size(476, 421);
            this.rightFormPanel.TabIndex = 1;
            // 
            // logoPlugin
            // 
            this.logoPlugin.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("logoPlugin.BackgroundImage")));
            this.logoPlugin.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.logoPlugin.Cursor = System.Windows.Forms.Cursors.No;
            this.logoPlugin.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logoPlugin.Location = new System.Drawing.Point(0, 0);
            this.logoPlugin.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.logoPlugin.Name = "logoPlugin";
            this.logoPlugin.Size = new System.Drawing.Size(492, 421);
            this.logoPlugin.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.logoPlugin.TabIndex = 0;
            this.logoPlugin.TabStop = false;
            this.logoPlugin.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // accountForm
            // 
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.White;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1237, 601);
            this.ControlBox = false;
            this.Controls.Add(this.formGeneralPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.Name = "accountForm";
            this.Padding = new System.Windows.Forms.Padding(4);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ARC Outlook Plugin";
            this.formGeneralPanel.ResumeLayout(false);
            this.leftLogoPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.logoPlugin)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel formGeneralPanel;
        private System.Windows.Forms.Panel rightFormPanel;
        private System.Windows.Forms.Panel leftLogoPanel;
        private System.Windows.Forms.PictureBox logoPlugin;
    }
}