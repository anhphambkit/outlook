using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ARC_Outlook_Plugin
{
    public partial class arc_addin
    {
        private void arc_addin_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_arc_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.showAccountForm(false);
        }
    }
}
