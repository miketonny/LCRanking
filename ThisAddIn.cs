using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace LCRanking
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.Application.WorkbookOpen  += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            ((AppEvents_Event)Application).NewWorkbook += new AppEvents_NewWorkbookEventHandler(Application_WB);
        }

        private void Application_WB(Excel.Workbook wb)
        {
            Style style1 = wb.Styles.Add("No1");
            Style style2 = wb.Styles.Add("No2");
            Style style3 = wb.Styles.Add("No3");
            Style style4 = wb.Styles.Add("No4");
            Style style5 = wb.Styles.Add("No5");
            style1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown);
            style2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            style3.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            style4.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            style5.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
