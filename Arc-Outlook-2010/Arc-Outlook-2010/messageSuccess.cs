using System;
using System.Windows.Forms;

namespace Arc_Outlook_2010
{
    public partial class messageSuccess : Form
    {
        public messageSuccess()
        {
            this.InitializeComponent();
        }

        private void animationSuccess_Tick(object sender, EventArgs e)
        {
            this.successImage.Enabled = false;
        }

        public void showMessage(string message)
        {
            this.messageLabel.Text = message;
            base.ShowDialog();
        }

        private void okBtnSuccessMessage_Click(object sender, EventArgs e)
        {
            base.Close();
        }
    }
}
