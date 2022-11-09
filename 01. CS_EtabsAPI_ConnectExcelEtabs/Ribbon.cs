using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ETABSv17;

namespace _01.CS_EtabsAPI_ConnectExcelEtabs
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void CheckStruture(object sender, RibbonControlEventArgs e)
        {
            
        }

        private void ClickEtabs(object sender, RibbonControlEventArgs e)
        {
            string ProgramPath = "C:\\Program Files (x86)\\Computers and Structures\\ETABS 17\\ETABS.exe";
            cOAPI myETABSObject = null;


        }
    }
}
