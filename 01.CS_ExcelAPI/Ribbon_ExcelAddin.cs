// ReSharper disable All
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ETABSv17;
using Microsoft.Office;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
// ReSharper disable All


namespace _01.CS_ExcelAPI
{
    public partial class RibbonExcelAddin
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
        public cOAPI etabModel = null;
        public cSapModel SapModel = null;
        EtabsClass etabsClass = new EtabsClass();
        public string comboName = "ENVESLS";
        public string comboUnits = "";

        //List<LoadCombination> LoadCombinationsList = new List<LoadCombination>();
        //List<JointReaction> JointReaction = new List<JointReaction>();
        //List<JointDisplacement> JointDiscplaList = new List<JointDisplacement>();

        private void BtnSelectEtabs_Click(object sender, RibbonControlEventArgs e)
        {
            etabsClass.SelectEtabs();
            etabModel = etabsClass.MyEtabsObject;
            SapModel = etabsClass.MySapModel;
        }

        private void BtnCheckStruc_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentWorksheet = Globals.ThisAddIn.GetActiveWorkSheet();
            int NumberPointNames = 1;
            string[] PointName = null;
            int NumberLevelNames = 1;
            string[] LevelName = null;




            int NumberResults = 1;
            string[] Obj = null;
            string[] Elm = null;
            string[] LoadCase = null;
            string[] StepType = null;
            double[] StepNum = null;
            double[] U1 = null;
            double[] U2 = null;
            double[] U3 = null;
            double[] R1 = null;
            double[] R2 = null;
            double[] R3 = null;
            SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int v = SapModel.Results.Setup.SetComboSelectedForOutput(comboName);
            SapModel.Story.GetNameList(ref NumberLevelNames, ref LevelName);
            for (int i = 0; i < LevelName.Length; i++)
            {
                SapModel.PointObj.GetNameListOnStory(LevelName[i], ref NumberPointNames, ref PointName);
                for (int j = 0; j < PointName.Length; j++)
                {
                    SapModel.Results.JointDispl(PointName[i], eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);
                    currentWorksheet.Cells[j * i + 1, 1] = LevelName[i];
                    currentWorksheet.Cells[j * i + 1, 2] = PointName[i];
                    currentWorksheet.Cells[j * i + 1, 3] = StepType[0];
                    currentWorksheet.Cells[j * i + 1, 4] = U1[0];
                    currentWorksheet.Cells[j * i + 1, 5] = U2[0];
                    currentWorksheet.Cells[j * i + 1, 6] = U3[0];
                    currentWorksheet.Cells[j * i + 1, 7] = R1[0];
                    currentWorksheet.Cells[j * i + 1, 8] = R2[0];
                    currentWorksheet.Cells[j * i + 1, 9] = R3[0];
                }

            }    

        }

        private void BtnReaction_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentWorksheet = Globals.ThisAddIn.GetActiveWorkSheet();
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


            SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int v = SapModel.Results.Setup.SetComboSelectedForOutput(comboName);
            SapModel.PointObj.GetNameListOnStory("Base", ref NumberNames, ref MyName);
            for (int i= 0; i < MyName.Length; i++)
            {
                SapModel.Results.JointReact(MyName[i],eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);
                currentWorksheet.Cells[2 * i + 1, 1] = MyName[i];
                currentWorksheet.Cells[2 * i + 1, 2] = StepType[0];
                currentWorksheet.Cells[2 * i + 1, 3] = F1[0];
                currentWorksheet.Cells[2 * i + 1, 4] = F2[0];
                currentWorksheet.Cells[2 * i + 1, 5] = F3[0];
                currentWorksheet.Cells[2 * i + 1, 6] = M1[0];
                currentWorksheet.Cells[2 * i + 1, 7] = M2[0];
                currentWorksheet.Cells[2 * i + 1, 8] = M3[0];

                //currentWorksheet.Cells[2 * i + 2, 1] = MyName[i];
                //currentWorksheet.Cells[2 * i + 2, 2] = StepType[1];
                //currentWorksheet.Cells[2 * i + 2, 3] = F1[1];
                //currentWorksheet.Cells[2 * i + 2, 4] = F2[1];
                //currentWorksheet.Cells[2 * i + 2, 5] = F3[1];
                //currentWorksheet.Cells[2 * i + 2, 6] = M1[1];
                //currentWorksheet.Cells[2 * i + 2, 7] = M2[1];
                //currentWorksheet.Cells[2 * i + 2, 8] = M3[1];

            }
        }

        private void BtnLoadData_Click(object sender, RibbonControlEventArgs e)
        {
            int NumberNames = 1;
            string[] MyName = null;
            List<eUnits> allUnits = new List<eUnits>();
                allUnits.Add(eUnits.N_m_C);
                allUnits.Add(eUnits.N_cm_C);
                allUnits.Add(eUnits.N_mm_C);
                allUnits.Add(eUnits.kN_m_C);
                allUnits.Add(eUnits.kN_cm_C);
                allUnits.Add(eUnits.kN_mm_C);
                allUnits.Add(eUnits.Ton_m_C);
                allUnits.Add(eUnits.Ton_cm_C);
                allUnits.Add(eUnits.Ton_mm_C);
                allUnits.Add(eUnits.kgf_m_C);
                allUnits.Add(eUnits.kgf_cm_C);
                allUnits.Add(eUnits.kgf_mm_C);
                allUnits.Add(eUnits.kip_ft_F);
                allUnits.Add(eUnits.kip_in_F);
                allUnits.Add(eUnits.lb_ft_F);
                allUnits.Add(eUnits.lb_in_F);

            for (int i = 0; i < allUnits.Count; i++)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = allUnits[i].ToString();
                comboBoxUnits.Items.Add(item);
            }
            string comboUnitsRibbon = comboBoxUnits.Text = comboBoxUnits.Items[3].Label;
            //comboUnits = comboUnitsRibbon;
            
            SapModel.RespCombo.GetNameList(ref NumberNames, ref MyName);
            for (int i = 0; i < MyName.Length; i++)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = MyName[i];
                comboBoxComboLoad.Items.Add(item);
            }
            string comboNameRibbon = comboBoxComboLoad.Text = comboBoxComboLoad.Items[0].Label;
            //comboName = comboNameRibbon;
        }
    }
}
