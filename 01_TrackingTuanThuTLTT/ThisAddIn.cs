using System;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace _01_TrackingTuanThuTLTT
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        public Worksheet GetActiveWorkSheet()
        {
            return (Worksheet)Application.ActiveSheet;
        }
        public Workbook GetActiveWorkBook()
        {
            return Application.ActiveWorkbook;
        }
        public Application GetActiveApp()
        {
            return Application.Application;
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
