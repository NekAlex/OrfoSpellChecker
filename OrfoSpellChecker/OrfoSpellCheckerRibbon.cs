namespace OrfoSpellChecker
{
    using System.Linq;
    using Microsoft.Office.Tools.Ribbon;
    using Excel = Microsoft.Office.Interop.Excel;
    public partial class OrfoSpellCheckerRibbon
    {
        private void OrfoSpellCheckerRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OSCCheckAll_Click(object sender, RibbonControlEventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var wss = excel.Worksheets;
            foreach (var ws in wss.OfType<Excel.Worksheet>())
            {
                SpellCheck.SpellCheckOnSheet(ws);
            }
        }
      
        private void OSCCheckCurrentTab_Click(object sender, RibbonControlEventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            SpellCheck.SpellCheckOnSheet(excel.ActiveSheet);
        }

        private void OSCAutoCheck_Click(object sender, RibbonControlEventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            if (OSCAutoCheck.Checked)
            {
                excel.Cells.Worksheet.Change += SpellCheck.Worksheet_Change;
            }
            else
            {
                excel.Cells.Worksheet.Change -= SpellCheck.Worksheet_Change;
            }

        }

        

        
    }
}
