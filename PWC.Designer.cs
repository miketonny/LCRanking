namespace LCRanking
{
    partial class PWC : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PWC()
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
            this.btnRun = this.Factory.CreateRibbonButton();
            this.btnRefresh = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.btnRun);
            this.group1.Items.Add(this.btnRefresh);
            this.group1.Label = "Loot Council";
            this.group1.Name = "group1";
            // 
            // btnRun
            // 
            this.btnRun.Label = "Run Parses";
            this.btnRun.Name = "btnRun";
            this.btnRun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRun_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Label = "Load Source";
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRefresh_Click);
            // 
            // PWC
            // 
            this.Name = "PWC";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PWC_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRun;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefresh;
    }

    partial class ThisRibbonCollection
    {
        internal PWC PWC
        {
            get { return this.GetRibbon<PWC>(); }
        }
    }
}
