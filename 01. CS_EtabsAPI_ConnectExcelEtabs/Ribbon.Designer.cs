
namespace _01.CS_EtabsAPI_ConnectExcelEtabs
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Etabs = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.CheckStructure = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Etabs.SuspendLayout();
            this.CheckStructure.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Etabs);
            this.tab1.Groups.Add(this.CheckStructure);
            this.tab1.Label = "ETABS Connect";
            this.tab1.Name = "tab1";
            // 
            // Etabs
            // 
            this.Etabs.DialogLauncher = ribbonDialogLauncherImpl1;
            this.Etabs.Items.Add(this.button1);
            this.Etabs.Items.Add(this.comboBox1);
            this.Etabs.Label = "ETABS";
            this.Etabs.Name = "Etabs";
            this.Etabs.Position = this.Factory.RibbonPosition.AfterOfficeId("");
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Chọn file Etabs";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ClickEtabs);
            // 
            // comboBox1
            // 
            this.comboBox1.Label = "Đơn vị";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Text = null;
            // 
            // CheckStructure
            // 
            this.CheckStructure.Items.Add(this.button2);
            this.CheckStructure.Label = "Kiểm tra ổn định kết cấu";
            this.CheckStructure.Name = "CheckStructure";
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Kiểm tra chuyển vị";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CheckStruture);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Etabs.ResumeLayout(false);
            this.Etabs.PerformLayout();
            this.CheckStructure.ResumeLayout(false);
            this.CheckStructure.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CheckStructure;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup Etabs;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
