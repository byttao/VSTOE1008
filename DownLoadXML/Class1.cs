using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace DownLoadXML
{
    public static class Class1
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rng"></param>
        /// <param name="str">获取的项目，如title、link</param>
        /// <param name="url">抓取地址</param>
        public static void DL(this Excel.Range rng,string str,string url)
        {
            XElement xml = XElement.Load(url);
            string txt = xml.Element("channel").Element("title").Value+":\r\n";
            var list =
                xml.Element("channel")
                    .Elements("item")
                    .Select((m, index1) => txt += index1.ToString() + "、" + m.Element(str).Value+"\r\n")
                    .ToList();
            rng.ColumnWidth = 100;
            rng.Value = txt;
            rng.Columns.AutoFit();
        } 
    }
}
