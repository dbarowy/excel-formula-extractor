using System;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace ExcelFormulaExtractor
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void extract_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application.Application;
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var graph = new Depends.DAG(wb, app, ignore_parse_errors: false, dagBuilt: new DateTime());
            System.Windows.Forms.MessageBox.Show("Dependence graph built");
        }
    }
}
