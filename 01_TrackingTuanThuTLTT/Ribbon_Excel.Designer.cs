
using System.ComponentModel;
using Microsoft.Office.Tools.Ribbon;

namespace _01_TrackingTuanThuTLTT
{
    partial class Ribbon_Excel : RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

        public Ribbon_Excel()
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
            this.button_daily = this.Factory.CreateRibbonButton();
            this.button_Monthly = this.Factory.CreateRibbonButton();
            this.button_loadData = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Tracking Tuan Thu ";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button_loadData);
            this.group1.Items.Add(this.button_daily);
            this.group1.Items.Add(this.button_Monthly);
            this.group1.Label = "Tracking";
            this.group1.Name = "group1";
            // 
            // button_daily
            // 
            this.button_daily.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_daily.Image = global::_01_TrackingTuanThuTLTT.Properties.Resources.Daily;
            this.button_daily.Label = "Hàng Ngày";
            this.button_daily.Name = "button_daily";
            this.button_daily.ShowImage = true;
            this.button_daily.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Daily_Click);
            // 
            // button_Monthly
            // 
            this.button_Monthly.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_Monthly.Image = global::_01_TrackingTuanThuTLTT.Properties.Resources.Monthly;
            this.button_Monthly.Label = "Hàng tháng";
            this.button_Monthly.Name = "button_Monthly";
            this.button_Monthly.ShowImage = true;
            this.button_Monthly.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Monthly_Click);
            // 
            // button_loadData
            // 
            this.button_loadData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_loadData.Label = "Load Dữ Liệu";
            this.button_loadData.Name = "button_loadData";
            this.button_loadData.ShowImage = true;
            this.button_loadData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LoadData_Click);
            // 
            // Ribbon_Excel
            // 
            this.Name = "Ribbon_Excel";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Excel_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal RibbonTab tab1;
        internal RibbonGroup group1;
        internal RibbonButton button_daily;
        internal RibbonButton button_Monthly;
        internal RibbonButton button_loadData;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_Excel Ribbon_Excel
        {
            get { return this.GetRibbon<Ribbon_Excel>(); }
        }
    }
}
