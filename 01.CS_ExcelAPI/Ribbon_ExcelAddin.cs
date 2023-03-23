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
        private void BtnSelectSafe_Click(object sender, RibbonControlEventArgs e)
        {
            etabsClass.SelectEtabs();
            etabModel = etabsClass.MyEtabsObject;
            SapModel = etabsClass.MySapModel;
        }

        private void BtnCheckStruc_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet currentWorksheet = Globals.ThisAddIn.GetActiveWorkSheet();
            
            int NumberPointNames = 1;
            string[] uniqueName = null;
            int StoryNumber = 1;
            string[] StoryName = null;
            double[] StoryHeight = null;
            double[] StoryElevation = null;
            bool[] IsMasterstory = null;
            string[] SimilarToStrory = null;
            bool[] SpiliceAbove = null;
            double[] SpliceHeight = null;



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

            
            List<string> levelName = new List<string>();
            SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int v = SapModel.Results.Setup.SetComboSelectedForOutput(comboName);
            SapModel.Story.GetStories(ref StoryNumber, ref StoryName, ref StoryElevation, ref StoryHeight,
                ref IsMasterstory, ref SimilarToStrory, ref SpiliceAbove, ref SpliceHeight);
            List<List<string>> storyNameList = new List<List<string>>();
            List<List<string>> pointNameList = new List<List<string>>();
            List<List<string>> steptypeList = new List<List<string>>();
            List<List<double>> U1list = new List<List<double>>();
            List<List<double>> U2list = new List<List<double>>();
            List<List<double>> U3list = new List<List<double>>();
            List<List<double>> R1list = new List<List<double>>();
            List<List<double>> R2list = new List<List<double>>();
            List<List<double>> R3list = new List<List<double>>();

            for (int i = 1; i < StoryName.Length; i++)
            {
                List<string> storyNameMemb = new List<string>();
                List<string> pointNameMemb = new List<string>();
                List<double> U1Member = new List<double>();
                List<double> U2Member = new List<double>();
                List<double> U3Member = new List<double>();
                List<double> R1Member = new List<double>();
                List<double> R2Member = new List<double>();
                List<double> R3Member = new List<double>();

                List<double> jointDisplacement = new List<double>(); 

                SapModel.PointObj.GetNameListOnStory(StoryName[i], ref NumberPointNames, ref uniqueName);
                //Lấy chuyển vị tất cả các point
                for (int j = 0; j < uniqueName.Length; j++)
                {
                    SapModel.Results.JointDispl(uniqueName[j], eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);
                    storyNameMemb.Add(StoryName[i]);
                    pointNameMemb.Add(uniqueName[j]);
                    U1Member.Add(U1[0]);
                    U2Member.Add(U2[0]);
                    U3Member.Add(U3[0]);
                    R1Member.Add(R1[0]);
                    R2Member.Add(R2[0]);
                    R3Member.Add(R3[0]);

                    jointDisplacement.Add(U1[0]);
                    jointDisplacement.Add(U1[1]);
                    jointDisplacement.Add(U2[0]);
                    jointDisplacement.Add(U2[1]);
                    jointDisplacement.Add(U3[0]);
                    jointDisplacement.Add(U3[1]);
                }

                currentWorksheet.Cells[i, 13] = StoryName[i];
                currentWorksheet.Cells[i, 14] = StoryElevation[i];
                currentWorksheet.Cells[i, 15] = jointDisplacement.Max();
                currentWorksheet.Cells[i, 16] = jointDisplacement.Min();

                storyNameList.Add(storyNameMemb);
                pointNameList.Add(pointNameMemb);
                U1list.Add(U1Member);
                U2list.Add(U1Member);
                U3list.Add(U1Member);
                R1list.Add(R1Member);
                R2list.Add(R2Member);
                R3list.Add(R3Member);

                //Lấy chuyển vị point lớn nhất tại mỗi tầng
                for (int k = 0; k < uniqueName.Length; k++)
                {

                }

            }
            
            List<string> resltstoName = storyNameList.SelectMany(i => i).ToList();
            List<string> resltpointName = pointNameList.SelectMany(i => i).ToList();
            List<double> reslU1 = U1list.SelectMany(i => i).ToList();
            List<double> reslU2 = U2list.SelectMany(i => i).ToList();
            List<double> reslU3 = U3list.SelectMany(i => i).ToList();
            List<double> reslR1 = R1list.SelectMany(i => i).ToList();
            List<double> reslR2 = R2list.SelectMany(i => i).ToList();
            List<double> reslR3 = R3list.SelectMany(i => i).ToList();
            for (int i = 0; i < resltstoName.Count(); i++)
            {
                currentWorksheet.Cells[i + 1, 1] = resltstoName[i];
                currentWorksheet.Cells[i + 1, 2] = resltpointName[i];
                currentWorksheet.Cells[i + 1, 3] = reslU1[i];
                currentWorksheet.Cells[i + 1, 4] = reslU2[i];
                currentWorksheet.Cells[i + 1, 5] = reslU3[i];
                currentWorksheet.Cells[i + 1, 6] = reslR1[i];
                currentWorksheet.Cells[i + 1, 7] = reslR2[i];
                currentWorksheet.Cells[i + 1, 8] = reslR3[i];
            }
        }

        private void BtnEtabsReaction_Click(object sender, RibbonControlEventArgs e)
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

        private void BtnAmVClick(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
