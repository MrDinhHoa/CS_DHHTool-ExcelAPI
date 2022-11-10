using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ETABSv17;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace _01.CS_EtabsAPI_ConnectExcelEtabs
{
    public partial class Ribbon
    {
        public cOAPI etabModel = null;
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            InitializeComponent();
        }
        public void ClickEtabs(object sender, RibbonControlEventArgs e)
        {
            
        }
        
        private void CheckStruture(object sender, RibbonControlEventArgs e)
        {
            
        }

        
    }
}
