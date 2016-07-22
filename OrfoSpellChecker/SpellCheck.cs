using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace OrfoSpellChecker
{
    public static class SpellCheck
    {
        public static void SpellChecker(Excel.Range target)
        {
            var app = Globals.ThisAddIn.Application.Application;
            string str = target.Text.ToString();
            if (app.CheckSpelling(str, Type.Missing, true) == false)
            {
                foreach (string tmp in ((string) str).Split(' '))
                {
                    if (app.CheckSpelling(tmp, Type.Missing, Type.Missing) == false)
                    {
                        if (target.Comment == null)
                        {
                            target.AddComment("Ошибка в слове " + tmp);
                        }
                        else
                        {
                            Excel.Characters c = target.Comment.Shape.TextFrame.Characters(Type.Missing, Type.Missing);
                            if (!c.Caption.Contains(tmp))
                            {
                                c.Caption = c.Caption + " " + tmp;
                            }
                        }
                        setFontColor(target, str.IndexOf(tmp) + 1, tmp.Length, 3);
                    }
                    else
                    {
                        setFontColor(target, str.IndexOf(tmp) + 1, tmp.Length, 0);
                    }
                }
            }
            else
            {
                if (target.Comment != null)
                {
                    if (target.Comment.Shape.AlternativeText.Contains("Ошибка в слове "))
                    {
                        setFontColor(target, str.IndexOf(str) + 1, str.Length, 0);
                        target.Comment.Delete();
                    }
                }
            }

        }

        private static void setFontColor(Excel.Range target, int startIdx, int strLen, int colorIndex)
        {
            target.Characters[startIdx, strLen].Font.ColorIndex = colorIndex;
        }
    }
}
