using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
namespace OrfoSpellChecker
{
    public partial class OrfoSpellCheckerRibbon
    {
        private void OrfoSpellCheckerRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OSCCheckAll_Click(object sender, RibbonControlEventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var wss = excel.Worksheets;
            var app = excel.Application;
            foreach (var ws in wss)
            {
                var sheet = ws as Excel.Worksheet;
                if (sheet != null)
                {
                    var range = sheet.UsedRange;
                    foreach (var cll in range)
                    {
                        var cell = cll as Excel.Range;
                        SpellCheck.SpellChecker(cell);
                    }
                }
            }
        }

        private void OSCCheckCurrentTab_Click(object sender, RibbonControlEventArgs e)
        {
            var excel = Globals.ThisAddIn.Application;
            var app = excel.Application;
            var sheet = app.ActiveSheet as Excel.Worksheet;
            if (sheet != null)
            {
                var range = sheet.UsedRange;
                foreach (var cll in range)
                {
                    var cell = cll as Excel.Range;
                    SpellCheck.SpellChecker(cell);
                }
            }
        }
    }
}
