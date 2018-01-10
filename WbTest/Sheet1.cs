using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace WbTest
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Range range = (Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet).Application.Selection.CurrentRegion;
            range.Interior.Color = System.Drawing.Color.LightBlue;
            //range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var range = (Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet).Range["A1","J10"];
            range.ClearFormats();
            var br1 = (Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet).Range["E4", "F7"];
            var br2 = (Globals.ThisWorkbook.ActiveSheet as Excel.Worksheet).Range["D5", "G6"];
            br1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            br2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        }
    }
}
