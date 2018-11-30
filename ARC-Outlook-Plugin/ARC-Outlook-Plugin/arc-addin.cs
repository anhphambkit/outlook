using Microsoft.Office.Tools.Ribbon;
using ARC_Outlook_Plugin.Properties;

namespace ARC_Outlook_Plugin
{
    public partial class arc_addin
    {
        private AboutArcPlugin _infoDialog;
        private void arc_addin_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_arc_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.showAccountForm(false);
        }

        private void syncBtn_Click(object sender, RibbonControlEventArgs e)
        {
            if (Settings.Default.token != null && Settings.Default.token != "")
            {
                Globals.ThisAddIn.StartSyncEmailNow();
            }
            else
            {
                Globals.ThisAddIn.showAccountForm(false);
            }
        }

        private void cleanDataBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.DeleteAllTemp();
            Globals.ThisAddIn.ShowSuccessMessage("Data temp removed success!");
        }

        private void inforBtn_Click(object sender, RibbonControlEventArgs e)
        {
            this._infoDialog = new AboutArcPlugin();
            this._infoDialog.showInfo();
        }

        private void arcWebBtn_Click(object sender, RibbonControlEventArgs e)
        {
            string url = Settings.Default.host;

            System.Diagnostics.Process.Start(url);
        }
    }
}
