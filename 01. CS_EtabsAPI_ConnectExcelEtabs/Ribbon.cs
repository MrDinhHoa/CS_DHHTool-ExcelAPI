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
        public cSapModel SapModel = null;
        EtabsClass etabsClass = new EtabsClass();
        List<LoadCombination> LoadCombinationsList = new List<LoadCombination>();
        List<JointReaction> JointReaction = new List<JointReaction>();
        List<JointDisplacement> JointDiscplaList = new List<JointDisplacement>();
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            InitializeComponent();
        }
        public void ClickEtabs(object sender, RibbonControlEventArgs e)
        {
            etabsClass.SelectEtabs();
            etabModel = etabsClass.MyETABSObject;
            SapModel = etabsClass.mySapModel;
        }
        private void reloadCombination(object sender, RibbonControlEventArgs e)
        {
            int NumberNames = 1;
            string[] MyName = null;

            SapModel.RespCombo.GetNameList(ref NumberNames, ref MyName);
            for (int i=0; i< MyName.Length; i++)
            {
                LoadCombination Lcomb = new LoadCombination();
                Lcomb.NumberNames = NumberNames;
                Lcomb.MyNames = MyName[i];
                LoadCombinationsList.Add(Lcomb);

            }    
        }
        private void ShowReactionLoad(object sender, RibbonControlEventArgs e)
        {
            string Name = "";
            int NumberResults = 1;
            string[] Obj = null;
            string[] Elm = null;
            string[] LoadCase = null;
            string[] StepType = null;
            double[] StepNum = null;
            double[] F1 = null;
            double[] F2 = null;
            double[] F3 =null;
            double[] M1 = null;
            double[] M2 = null;
            double[] M3 = null;
            SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int v = SapModel.Results.Setup.SetComboSelectedForOutput("ENVESLS");
            SapModel.Results.JointReact("37", eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);
            JointReaction jReactions = new JointReaction();
            jReactions.Name = Name;
            jReactions.LoadCase = LoadCase [0];
            jReactions.F1 = F1[0];
            jReactions.F1 = F2[0];
            jReactions.F1 = F3[0];
            jReactions.F1 = M1[0];
            jReactions.F1 = M2[0];
            jReactions.F1 = M3[0];
            JointReaction.Add(jReactions);
        }
        private void ShowDisplacement(object sender, RibbonControlEventArgs e)
        {
            string Name = "";
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
            SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();
            int v = SapModel.Results.Setup.SetComboSelectedForOutput("ENVESLS");
            SapModel.Results.JointDispl("37", eItemTypeElm.Element, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref F1, ref F2, ref F3, ref M1, ref M2, ref M3);
            JointReaction jReactions = new JointReaction();
            jReactions.Name = Name;
            jReactions.LoadCase = LoadCase[0];
            jReactions.F1 = F1[0];
            jReactions.F1 = F2[0];
            jReactions.F1 = F3[0];
            jReactions.F1 = M1[0];
            jReactions.F1 = M2[0];
            jReactions.F1 = M3[0];
            JointReaction.Add(jReactions);
        }
    }
}
