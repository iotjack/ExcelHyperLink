using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace HyperLink
{
    //save last url used in settings
    //https://docs.microsoft.com/en-us/previous-versions/aa730869(v=vs.80)
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.button1.Click += new RibbonControlEventHandler(this.button1_Click);
            editBox1.Text = Properties.Settings.Default.lastURL;

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddHyperLink(editBox1.Text);
            Properties.Settings.Default.lastURL = editBox1.Text.Trim();  
            Properties.Settings.Default.Save();
        }
    }
}
