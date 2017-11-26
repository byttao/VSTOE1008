using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using DownLoadXML;
using Excel = Microsoft.Office.Interop.Excel;

namespace VSTO作业20171121
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (editBox1.Text == "" || editBox2.Text == "")
                MessageBox.Show("请填写完整后重试！");
            else
            {
                Excel.Range rng = Globals.ThisAddIn.Application.ActiveCell;
                rng.DL(editBox1.Text, editBox2.Text);
            }
        }
    }
}
