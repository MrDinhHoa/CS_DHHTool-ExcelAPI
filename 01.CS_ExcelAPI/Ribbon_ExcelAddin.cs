using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ETABSv17;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Worksheet = Microsoft.Office.Tools.Excel.Worksheet;

namespace _01.CS_ExcelAPI
{
    public partial class RibbonExcelAddin
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
        public cOAPI etabModel = null;
        public cSapModel SapModel = null;
        
        // ReSharper disable once ArrangeTypeMemberModifiers
        EtabsClass etabsClass = new EtabsClass();
        List<LoadCombination> LoadCombinationsList = new List<LoadCombination>();
        List<JointReaction> JointReaction = new List<JointReaction>();
        List<JointDisplacement> JointDiscplaList = new List<JointDisplacement>();

        private void BtnSelectEtabs_Click(object sender, RibbonControlEventArgs e)
        {
            etabsClass.SelectEtabs();
            etabModel = etabsClass.MyEtabsObject;
            SapModel = etabsClass.MySapModel;
        }

        private void BtnCheckStruc_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void BtnReaction_Click(object sender, RibbonControlEventArgs e)
        {
            //string Name = "";
            Excel.Application xlApp = new Application();
            Excel.Workbook xlwWorkbook = xlApp.ActiveWorkbook;
            Excel.Worksheet xlWorksheet = xlwWorkbook.ActiveSheet;


            int NumberResults = 1;
            string[] Obj = null;
            string[] Elm = null;
            string[] LoadCase = null;
            string[] StepType = null;
            double[] StepNum = null;
            double[] F1 = null;
            double[] F2 = null;
            double[] F3 = null;
            double[] M1 = null;
            double[] M2 = null;
            double[] M3 = null;


            int NumberNames = 1;
            string[] MyName = null;
            //double[] X = null;
            //double[] Y = null;
            //double[] Z = null;

            SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int v = SapModel.Results.Setup.SetComboSelectedForOutput("ENVESLS");
            //SapModel.Results.JointReact("ALL", eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);
            JointReaction jReactions = new JointReaction();
            //jReactions.Name = Name;
            //jReactions.LoadCase = "ENVESLS";
            //jReactions.F1 = F1[0];
            //jReactions.F2 = F2[0];
            //jReactions.F3 = F3[0];
            //jReactions.M1 = M1[0];
            //jReactions.M2 = M2[0];
            //jReactions.M3 = M3[0];
            //JointReaction.Add(jReactions);
            
            SapModel.PointObj.GetNameListOnStory("Base", ref NumberNames, ref MyName);
            for (int i= 0; i < MyName.Length; i++)
            {
                SapModel.Results.JointReact(MyName[i],eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);
                xlWorksheet.Cells[i + 1, 1] = F1[0];
                xlWorksheet.Cells[i + 1, 1] = F2[0];
                xlWorksheet.Cells[i + 1, 1] = F3[0];
                xlWorksheet.Cells[i + 1, 1] = M1[0];
                xlWorksheet.Cells[i + 1, 1] = M2[0];
                xlWorksheet.Cells[i + 1, 1] = M3[0];

            }
        }
    }
}
