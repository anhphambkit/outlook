namespace Arc_Outlook_2010
{
    partial class messageSuccess
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(messageSuccess));
            this.imageLayoutMessage = new System.Windows.Forms.Panel();
            this.successImage = new System.Windows.Forms.PictureBox();
            this.contentMessageLayout = new System.Windows.Forms.Panel();
            this.messageLabel = new System.Windows.Forms.Label();
            this.actionMessageLayout = new System.Windows.Forms.Panel();
            this.okBtnSuccessMessage = new System.Windows.Forms.Button();
            this.animationSuccess = new System.Windows.Forms.Timer(this.components);
            this.imageLayoutMessage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.successImage)).BeginInit();
            this.contentMessageLayout.SuspendLayout();
            this.actionMessageLayout.SuspendLayout();
            this.SuspendLayout();
            // 
            // imageLayoutMessage
            // 
            this.imageLayoutMessage.BackColor = System.Drawing.Color.White;
            this.imageLayoutMessage.Controls.Add(this.successImage);
            this.imageLayoutMessage.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.imageLayoutMessage.Dock = System.Windows.Forms.DockStyle.Top;
            this.imageLayoutMessage.Font = new System.Drawing.Font("Microsoft Tai Le", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.imageLayoutMessage.ForeColor = System.Drawing.Color.DimGray;
            this.imageLayoutMessage.Location = new System.Drawing.Point(0, 0);
            this.imageLayoutMessage.Name = "imageLayoutMessage";
            this.imageLayoutMessage.Size = new System.Drawing.Size(318, 280);
            this.imageLayoutMessage.TabIndex = 14;
            // 
            // successImage
            // 
            this.successImage.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.successImage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.successImage.Image = ((System.Drawing.Image)(resources.GetObject("successImage.Image")));
            this.successImage.Location = new System.Drawing.Point(0, 0);
            this.successImage.Margin = new System.Windows.Forms.Padding(2);
            this.successImage.Name = "successImage";
            this.successImage.Size = new System.Drawing.Size(318, 280);
            this.successImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.successImage.TabIndex = 0;
            this.successImage.TabStop = false;
            // 
            // contentMessageLayout
            // 
            this.contentMessageLayout.Controls.Add(this.messageLabel);
            this.contentMessageLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.contentMessageLayout.Location = new System.Drawing.Point(0, 280);
            this.contentMessageLayout.Margin = new System.Windows.Forms.Padding(2);
            this.contentMessageLayout.Name = "contentMessageLayout";
            this.contentMessageLayout.Size = new System.Drawing.Size(318, 62);
            this.contentMessageLayout.TabIndex = 15;
            // 
            // messageLabel
            // 
            this.messageLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.messageLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.messageLabel.ForeColor = System.Drawing.SystemColors.AppWorkspace;
            this.messageLabel.Location = new System.Drawing.Point(0, 0);
            this.messageLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.messageLabel.Name = "messageLabel";
            this.messageLabel.Size = new System.Drawing.Size(318, 62);
            this.messageLabel.TabIndex = 14;
            this.messageLabel.Text = "Login Success!";
            this.messageLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // actionMessageLayout
            // 
            this.actionMessageLayout.Controls.Add(this.okBtnSuccessMessage);
            this.actionMessageLayout.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.actionMessageLayout.Location = new System.Drawing.Point(0, 346);
            this.actionMessageLayout.Margin = new System.Windows.Forms.Padding(2);
            this.actionMessageLayout.Name = "actionMessageLayout";
            this.actionMessageLayout.Size = new System.Drawing.Size(318, 70);
            this.actionMessageLayout.TabIndex = 16;
            // 
            // okBtnSuccessMessage
            // 
            this.okBtnSuccessMessage.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(119)))), ((int)(((byte)(180)))), ((int)(((byte)(63)))));
            this.okBtnSuccessMessage.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.okBtnSuccessMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.okBtnSuccessMessage.ForeColor = System.Drawing.Color.White;
            this.okBtnSuccessMessage.Location = new System.Drawing.Point(119, 12);
            this.okBtnSuccessMessage.Margin = new System.Windows.Forms.Padding(2);
            this.okBtnSuccessMessage.Name = "okBtnSuccessMessage";
            this.okBtnSuccessMessage.Size = new System.Drawing.Size(79, 37);
            this.okBtnSuccessMessage.TabIndex = 0;
            this.okBtnSuccessMessage.Text = "OK";
            this.okBtnSuccessMessage.UseVisualStyleBackColor = false;
            this.okBtnSuccessMessage.Click += new System.EventHandler(this.okBtnSuccessMessage_Click);
            // 
            // animationSuccess
            // 
            this.animationSuccess.Enabled = true;
            this.animationSuccess.Interval = 4000;
            // 
            // messageSuccess
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(318, 416);
            this.ControlBox = false;
            this.Controls.Add(this.actionMessageLayout);
            this.Controls.Add(this.contentMessageLayout);
            this.Controls.Add(this.imageLayoutMessage);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "messageSuccess";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.imageLayoutMessage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.successImage)).EndInit();
            this.contentMessageLayout.ResumeLayout(false);
            this.actionMessageLayout.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel imageLayoutMessage;
        private System.Windows.Forms.PictureBox successImage;
        private System.Windows.Forms.Panel contentMessageLayout;
        private System.Windows.Forms.Label messageLabel;
        private System.Windows.Forms.Panel actionMessageLayout;
        private System.Windows.Forms.Button okBtnSuccessMessage;
        private System.Windows.Forms.Timer animationSuccess;
    }
}