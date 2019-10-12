using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace HyperLink
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        internal void AddHyperLink(string prefix)
        {
            Excel.Worksheet worksheet = this.Application.ActiveWorkbook.ActiveSheet;
            Excel.Range r = worksheet.Cells.Application.Selection;
            foreach (object cell in r.Cells)
            {
                worksheet.Hyperlinks.Add(cell, prefix + ((Excel.Range)cell).Value2.ToString(), Type.Missing, "", "");
                //((Excel.Range)cell).Value2 = "prefix";
            }
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
