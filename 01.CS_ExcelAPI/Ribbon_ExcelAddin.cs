// ReSharper disable All
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ETABSv17;
using Microsoft.Office;
using Microsoft.Office.Core;    
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using eItemType = CSiAPIv1.eItemType;
using Excel = Microsoft.Office.Interop.Excel;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using System.Security.Cryptography.X509Certificates;
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

            //Get all Load Combo in Model
            int NumbernameCombo = 1;
            string[] ComboName = null;
            SapModel.RespCombo.GetNameList(ref NumbernameCombo, ref ComboName);
            string[] SLSComboName = Array.FindAll(ComboName, x => x.StartsWith("SLS"));

            //Get All Stories in Model
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

            //Get Displacement Max by Level And Load Combo
            SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            var jointDisplacements = SLSComboName.Select
            (slsComboname =>
            {
                SapModel.Results.Setup.SetComboSelectedForOutput(slsComboname);
                List<JointDisplacement> jointDisplacement = StoryName.AsParallel().Select
                (storyName =>
                {
                    SapModel.PointObj.GetNameListOnStory(storyName, ref NumberPointNames, ref uniqueName);

                    List<JointDisplacement> jdisps = uniqueName.AsParallel().Select(unique =>
                    {
                        Stopwatch jointDisplStopwatch = Stopwatch.StartNew();
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
                        SapModel.Results.JointDispl(unique, eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum,
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
                    }).ToList();

                    double maxUx = jdisps.Max(k => k.Ux);
                    double minUy = jdisps.Min(k => k.Uy);

                    JointDisplacement jdisp = jdisps.First();

                    return new JointDisplacement()
                    {
                        Level = jdisp.Level,
                        Ux = maxUx,
                        Uy = minUy
                    };
                }
                ).ToList();
            }
            ).ToArray();

            #region Get Displacement Not Load Combo
            //Stopwatch stopwatch = new Stopwatch();
            //stopwatch.Start();

            //JointDisplacement[] jointDisplacement = StoryName.AsParallel().Select
            //(storyName =>
            //{
            //    Stopwatch getNameListStopwatch = Stopwatch.StartNew();
            //    SapModel.PointObj.GetNameListOnStory(storyName, ref NumberPointNames, ref uniqueName);
            //    getNameListStopwatch.Stop();
            //    Debug.WriteLine($"GetNameListOnStory: {stopwatch.ElapsedMilliseconds} ms");

            //    List<JointDisplacement> jdisps = uniqueName.AsParallel().Select(unique =>
            //    {
            //        Stopwatch jointDisplStopwatch = Stopwatch.StartNew();
            //        int NumberResults = 1;
            //        string[] Obj = null;
            //        string[] Elm = null;
            //        string[] LoadCase = null;
            //        string[] StepType = null;
            //        double[] StepNum = null;
            //        double[] U1 = null;
            //        double[] U2 = null;
            //        double[] U3 = null;
            //        double[] R1 = null;
            //        double[] R2 = null;
            //        double[] R3 = null;
            //        SapModel.Results.JointDispl(unique, eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum,
            //                                     ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);
            //        jointDisplStopwatch.Stop();
            //        Debug.WriteLine($"JointDispl: {stopwatch.ElapsedMilliseconds} ms");

            //        return new JointDisplacement()
            //        {
            //            Level = storyName,
            //            Name = unique,
            //            LoadCase = comboName,
            //            Ux = U1[0],
            //            Uy = U2[0],
            //            Uz = U3[0],
            //            Rx = R1[0],
            //            Ry = R2[0],
            //            Rz = R3[0]
            //        };
            //    }).ToList();

            //    double maxUx = jdisps.Max(k => k.Ux);
            //    double minUy = jdisps.Min(k => k.Uy);

            //    JointDisplacement jdisp = jdisps.First();

            //    return new JointDisplacement()
            //    {
            //        Level = jdisp.Level,
            //        Ux = maxUx,
            //        Uy = minUy
            //    };
            //}
            //).ToArray();
            #endregion

            #region Write To Excel
            Parallel.For(0, jointDisplacements.Length, i =>
            {
                var jdp = jointDisplacements[i];
                currentWorksheet.Cells[i + 1, 3] = jdp.Level;
                currentWorksheet.Cells[i + 1, 4] = comboName;
                currentWorksheet.Cells[i + 1, 6] = jdp.Ux;
            });
            stopwatch.Stop();
            Debug.WriteLine($"Whole process execution time: {stopwatch.ElapsedMilliseconds} ms");
            #endregion
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
