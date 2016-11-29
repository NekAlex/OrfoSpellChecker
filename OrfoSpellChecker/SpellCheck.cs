namespace OrfoSpellChecker
{
    using System;
    using System.Linq;
    using Excel = Microsoft.Office.Interop.Excel;
    public static class SpellCheck
    {
        public static void SpellChecker(Excel.Range target)
        {
            var app = Globals.ThisAddIn.Application.Application;
            string str = target.Text.ToString();
            if (app.CheckSpelling(str, Type.Missing, true) == false)
            {
                foreach (var tmp in str.Split(' '))
                {
                    if (app.CheckSpelling(tmp, Type.Missing, Type.Missing) == false)
                    {
                        if (target.Comment == null)
                        {
                            target.AddComment("Ошибка в слове " + tmp);
                        }
                        else
                        {
                            var c = target.Comment.Shape.TextFrame.Characters(Type.Missing, Type.Missing);
                            if (!c.Caption.Contains(tmp))
                            {
                                c.Caption = c.Caption + " " + tmp;
                            }
                        }
                        SetFontColor(target, str.IndexOf(tmp, StringComparison.Ordinal) + 1, tmp.Length, 3);
                    }
                    else
                    {
                        SetFontColor(target, str.IndexOf(tmp, StringComparison.Ordinal) + 1, tmp.Length, 0);
                    }
                }
            }
            else
            {
                if (target.Comment != null)
                {
                    if (target.Comment.Shape.AlternativeText.Contains("Ошибка в слове "))
                    {
                        SetFontColor(target, str.IndexOf(str, StringComparison.Ordinal) + 1, str.Length, 0);
                        target.Comment.Delete();
                    }
                }
            }

        }

        private static void SetFontColor(Excel.Range target, int startIdx, int strLen, int colorIndex)
        {
            target.Characters[startIdx, strLen].Font.ColorIndex = colorIndex;
        }

        public static void Worksheet_Change(Excel.Range target)
        {
            SpellChecker(target);
        }
        public static void SpellCheckOnSheet(Excel.Worksheet sheet)
        {
            if (sheet == null) return;
            var range = sheet.UsedRange;
            foreach (var cell in range.OfType<Excel.Range>())
            {
                Worksheet_Change(cell);
            }
        }
    }
}
