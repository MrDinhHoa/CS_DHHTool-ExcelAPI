// ReSharper disable All
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ETABSv17;
using Microsoft.Office;
using Microsoft.Office.Core;    
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using eItemType = CSiAPIv1.eItemType;
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
        public CSiAPIv1.cSapModel MySapModel = null;
        EtabsClass etabsClass = new EtabsClass();
        SapClass sapClass = new SapClass();
        public string comboName = "ENVESLS";
        public string comboUnits = "";

        //List<LoadCombination> LoadCombinationsList = new List<LoadCombination>();
        //List<JointReaction> JointReaction = new List<JointReaction>();
        //List<JointDisplacement> JointDiscplaList = new List<JointDisplacement>();

        private void BtnSelectEtabs_Click(object sender, RibbonControlEventArgs e)
        {
            etabsClass.SelectEtabs();
            etabModel = (cOAPI)etabsClass.MyEtabsObject;
            SapModel = (cSapModel)etabsClass.MySapModel;
        }

        private void BtnSAP_Click(object sender, RibbonControlEventArgs e)
        {
            sapClass.SelectSAP();
            MySapModel = sapClass.mySapModel;

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
            SapModel.Story.GetStories(ref StoryNumber, ref StoryName, ref StoryElevation, ref StoryHeight,
                ref IsMasterstory, ref SimilarToStrory, ref SpiliceAbove, ref SpliceHeight);


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

            List<List<string>> storyNameList = new List<List<string>>();
            List<List<string>> pointNameList = new List<List<string>>();
            List<List<string>> steptypeList = new List<List<string>>();
            List<List<double>> U1list = new List<List<double>>();
            List<List<double>> U2list = new List<List<double>>();
            List<List<double>> U3list = new List<List<double>>();
            List<List<double>> R1list = new List<List<double>>();
            List<List<double>> R2list = new List<List<double>>();
            List<List<double>> R3list = new List<List<double>>();
            // Length = 6
            var jointDisplacement  = StoryName.Select((storyName) =>
            {
                SapModel.PointObj.GetNameListOnStory(storyName, ref NumberPointNames, ref uniqueName);
                var jdisps = uniqueName.Select((unique) =>
                {
                    SapModel.Results.JointDispl(unique, eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm,
                        ref LoadCase, ref StepType, ref StepNum,
                        ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);
                    return new JointDisplacement()
                    {
                        Level = storyName,
                        Name = unique,
                        LoadCase = comboName,
                        Ux = U1[0],
                        Uy = U2[0],
                        Uz = U3[0],
                        Rx = R1[0],
                        Ry = R2[0],
                        Rz = R3[0]

                    };
                });
                // jdisps.Max((k) => k.Ux) returm phan tu co UX lon nhat
                // [1 1 2 3 4  6 6]
                // jdisps.Max((k) => k.Ux)= 6
                // [6 6]
                var jdisp = jdisps.First();
                // var jdispMaxUx = jdisps.First((j) => j.Ux == jdisps.Max((k) => k.Ux));
                


                return new JointDisplacement()
                {
                    Level = jdisp.Level,
                    Ux = jdisps.Max((k) => k.Ux),
                    Uy = jdisps.Min((k) => k.Uy),
                };
            }).ToArray();
            for (var i = 0;i < jointDisplacement.Count(); i++)
            {
                var jdp = jointDisplacement[i];
                currentWorksheet.Cells[i+1, 3] = jdp.Level;
                currentWorksheet.Cells[i+1, 4] = comboName;
                currentWorksheet.Cells[i+1, 6] = jdp.Ux;
            }


            //for (int i = 1; i < StoryName.Length; i++)
            //{
            //    List<string> storyNameMemb = new List<string>();
            //    List<string> pointNameMemb = new List<string>();
            //    List<double> U1Member = new List<double>();
            //    List<double> U2Member = new List<double>();
            //    List<double> U3Member = new List<double>();
            //    List<double> R1Member = new List<double>();
            //    List<double> R2Member = new List<double>();
            //    List<double> R3Member = new List<double>();
            //    //List<JointDisplacement> jointDisplacement = new List<JointDisplacement>(); 
            //    SapModel.PointObj.GetNameListOnStory(StoryName[i], ref NumberPointNames, ref uniqueName);
            //    //Lấy chuyển vị tất cả các point
            //    for (int j = 0; j < uniqueName.Length; j++)
            //    {
            //        SapModel.Results.JointDispl(uniqueName[j], eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, 
            //            ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);

            //        {
            //            for (int k = 0; k <=1; k++)
            //            {
            //                JointDisplacement jdisp = new JointDisplacement();
            //                jdisp.Level = StoryName[i];
            //                jdisp.Name = uniqueName[j];
            //                jdisp.LoadCase = comboName;
            //                jdisp.Ux = U1[k];
            //                jdisp.Uy = U2[k];
            //                jdisp.Uz = U3[k];
            //                jdisp.Rx = R1[k];
            //                jdisp.Ry = R2[k];
            //                jdisp.Rz = R3[k];
            //                jointDisplacement.Add(jdisp);
            //            }


            //        }
            //    }

            //    currentWorksheet.Cells[i, 3] = StoryName[i];
            //    currentWorksheet.Cells[i, 4] = comboName; 
            //    currentWorksheet.Cells[i, 5] = StoryElevation[i];
            //    currentWorksheet.Cells[i, 6] = jointDisplacement.Max(x => x.Ux);
            //    currentWorksheet.Cells[i, 7] = jointDisplacement.Max(x => x.Uy);
            //    currentWorksheet.Cells[i, 8] = jointDisplacement.Min(x => x.Ux);
            //    currentWorksheet.Cells[i, 9] = jointDisplacement.Min(x => x.Uy);
            //    currentWorksheet.Cells[i, 10] = Math.Max(Math.Abs(jointDisplacement.Max(x => x.Ux)), Math.Abs(jointDisplacement.Min(x => x.Ux)));
            //    currentWorksheet.Cells[i, 11] = Math.Max(Math.Abs(jointDisplacement.Max(x => x.Uy)), Math.Abs(jointDisplacement.Min(x => x.Uy)));

            //}

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
            int I = MySapModel.FrameObj.SetSelected("ALL", true, eItemType.Group);
            MySapModel.SelectObj.ClearSelection();
            
            int NumberNames = 1;
            int[] Object = null;
            string[] MyName = null;
            MySapModel.SelectObj.GetSelected(ref NumberNames, ref Object, ref MyName);


        }

    }
}
