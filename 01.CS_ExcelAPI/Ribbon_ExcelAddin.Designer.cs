
using Microsoft.Office.Tools.Ribbon;

namespace _01.CS_ExcelAPI
{
    partial class RibbonExcelAddin : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonExcelAddin()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonExcelAddin));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.comboBoxUnits = this.Factory.CreateRibbonComboBox();
            this.comboBoxComboLoad = this.Factory.CreateRibbonComboBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.BtnSelectEtabs = this.Factory.CreateRibbonButton();
            this.BtnLoadData = this.Factory.CreateRibbonButton();
            this.BtnCheckStruc = this.Factory.CreateRibbonButton();
            this.BtnReaction = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "ETABS Connect";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.BtnLoadData);
            this.group2.Items.Add(this.comboBoxUnits);
            this.group2.Items.Add(this.comboBoxComboLoad);
            this.group2.Label = "Dữ liệu ETABS";
            this.group2.Name = "group2";
            // 
            // comboBoxUnits
            // 
            this.comboBoxUnits.Label = "Đơn vị";
            this.comboBoxUnits.MaxLength = 40;
            this.comboBoxUnits.Name = "comboBoxUnits";
            this.comboBoxUnits.Text = null;
            this.comboBoxUnits.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCheckStruc_Click);
            // 
            // comboBoxComboLoad
            // 
            this.comboBoxComboLoad.Label = "Tổ hợp";
            this.comboBoxComboLoad.MaxLength = 40;
            this.comboBoxComboLoad.Name = "comboBoxComboLoad";
            this.comboBoxComboLoad.Tag = "";
            this.comboBoxComboLoad.Text = null;
            // 
            // group4
            // 
            this.group4.Items.Add(this.BtnCheckStruc);
            this.group4.Items.Add(this.BtnReaction);
            this.group4.Label = "Kiểm tra ổn định kết cấu";
            this.group4.Name = "group4";
            // 
            // BtnSelectEtabs
            // 
            this.BtnSelectEtabs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnSelectEtabs.Image = ((System.Drawing.Image)(resources.GetObject("BtnSelectEtabs.Image")));
            this.BtnSelectEtabs.Label = "Kết nối ETABS";
            this.BtnSelectEtabs.Name = "BtnSelectEtabs";
            this.BtnSelectEtabs.ShowImage = true;
            this.BtnSelectEtabs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSelectEtabs_Click);
            // 
            // BtnLoadData
            // 
            this.BtnLoadData.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadData.Image")));
            this.BtnLoadData.Label = "Load dữ liệu";
            this.BtnLoadData.Name = "BtnLoadData";
            this.BtnLoadData.ShowImage = true;
            this.BtnLoadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadData_Click);
            // 
            // BtnCheckStruc
            // 
            this.BtnCheckStruc.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnCheckStruc.Image = ((System.Drawing.Image)(resources.GetObject("BtnCheckStruc.Image")));
            this.BtnCheckStruc.Label = "Kiểm tra chuyển vị";
            this.BtnCheckStruc.Name = "BtnCheckStruc";
            this.BtnCheckStruc.ShowImage = true;
            this.BtnCheckStruc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCheckStruc_Click);
            // 
            // BtnReaction
            // 
            this.BtnReaction.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnReaction.Image = ((System.Drawing.Image)(resources.GetObject("BtnReaction.Image")));
            this.BtnReaction.Label = "Phản lực chân cột";
            this.BtnReaction.Name = "BtnReaction";
            this.BtnReaction.ShowImage = true;
            this.BtnReaction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnReaction_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.BtnSelectEtabs);
            this.group1.Label = "ETABS";
            this.group1.Name = "group1";
            // 
            // RibbonExcelAddin
            // 
            this.Name = "RibbonExcelAddin";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadData;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCheckStruc;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxUnits;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnReaction;
        public Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxComboLoad;
        internal RibbonButton BtnSelectEtabs;
        internal RibbonGroup group1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonExcelAddin Ribbon
        {
            get { return this.GetRibbon<RibbonExcelAddin>(); }
        }
    }
}
