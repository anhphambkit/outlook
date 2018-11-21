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
            this.formGeneralPanel = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // formGeneralPanel
            // 
            this.formGeneralPanel.AutoSize = true;
            this.formGeneralPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.formGeneralPanel.BackColor = System.Drawing.Color.Transparent;
            this.formGeneralPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.formGeneralPanel.Location = new System.Drawing.Point(4, 4);
            this.formGeneralPanel.Margin = new System.Windows.Forms.Padding(33, 32, 33, 32);
            this.formGeneralPanel.Name = "formGeneralPanel";
            this.formGeneralPanel.Padding = new System.Windows.Forms.Padding(100, 86, 100, 86);
            this.formGeneralPanel.Size = new System.Drawing.Size(1229, 593);
            this.formGeneralPanel.TabIndex = 11;
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
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel formGeneralPanel;
    }
}