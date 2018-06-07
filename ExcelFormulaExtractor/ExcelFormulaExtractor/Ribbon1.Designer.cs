namespace ExcelFormulaExtractor
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.extract = this.Factory.CreateRibbonButton();
            this.ExtractThis = this.Factory.CreateRibbonButton();
            this.ExtractToFPCore = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.extract);
            this.group1.Items.Add(this.ExtractThis);
            this.group1.Items.Add(this.ExtractToFPCore);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // extract
            // 
            this.extract.Label = "Extract All";
            this.extract.Name = "extract";
            this.extract.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.extract_Click);
            // 
            // ExtractThis
            // 
            this.ExtractThis.Label = "Extract This";
            this.ExtractThis.Name = "ExtractThis";
            this.ExtractThis.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExtractThis_Click);
            // 
            // ExtractToFPCore
            // 
            this.ExtractToFPCore.Label = "Extract All to FPCore";
            this.ExtractToFPCore.Name = "ExtractToFPCore";
            this.ExtractToFPCore.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExtractToFPCore_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton extract;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExtractThis;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExtractToFPCore;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
