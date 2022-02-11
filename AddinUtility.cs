using Microsoft.Office.Interop.Word;
using System;
using System.Text;
using System.Text.RegularExpressions;

namespace PaperWriting
{
    public class AddinUtility
    {
        private const string bracketRegex = @"((?<!\\)\{.*?(?<!\\)\}|(?<!\\)[\[\]]|(?<!\\)#)";

        private static Selection selection = Globals.ThisAddIn.Application.Selection;
        private static Document document = Globals.ThisAddIn.Application.ActiveDocument;
        private static UndoRecord undoRecord = Globals.ThisAddIn.Application.UndoRecord;

        public static Range InsertOMath(Range range = null)
        {
            if (range == null) range = selection.Range;
            return range.OMaths.Add(range);
        }

        public static Range InsertField(string code, Range range = null)
        {
            if (range == null) range = selection.Range;
            return range.Fields.Add(range, WdFieldType.wdFieldEmpty, code, false).Result;
        }

        public static Range InsertContent(string content, string style=null, Range range = null)
        {
            undoRecord.StartCustomRecord("论文辅助-插入描述");
            if (range == null) range = selection.Range;
            try
            {
                range.set_Style(style);
            }catch (System.Runtime.InteropServices.COMException) { }
            Range bookmarkRange = document.Range();
            var contentTextArray = Regex.Split(content, bracketRegex);
            foreach (string contentText in contentTextArray)
            {
                if (contentText == "") continue;
                if (contentText.StartsWith("{") && contentText.EndsWith("}"))
                {
                    range=InsertField(contentText.Substring(1, contentText.Length - 2), range);
                    range.Move(Count: range.End - range.Start);
                    continue;
                }
                if (contentText == "[")
                {
                    bookmarkRange.Start = range.Start;
                    continue;
                }
                if (contentText == "]")
                {
                    bookmarkRange.End = range.End;
                    continue;
                }
                if (contentText == "#")
                {
                    selection.SetRange(range.Start,range.Start);
                    continue;
                }
                range.InsertAfter(contentText.Replace(@"\#","#").Replace(@"\{","{").Replace(@"\}","}").Replace(@"\[","[").Replace(@"\]","]"));
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
            }
            document.Bookmarks.Add(GenerateBookmarkName(), bookmarkRange);
            undoRecord.EndCustomRecord();
            return range;
        }

        public static string GenerateBookmarkName()
        {
            char[] constant =
            {
                '0','1','2','3','4','5','6','7','8','9',
                'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z',
                'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'
            };
            StringBuilder bookmarkName;
            Random random = new Random();
            do
            {
                bookmarkName = new StringBuilder(Properties.Settings.Default.BookmarkPrefix);
                for (int i = 0; i < 8; i++)
                {
                    bookmarkName.Append(constant[random.Next(constant.Length)]);
                }
            } while (document.Bookmarks.Exists(bookmarkName.ToString()));
            return bookmarkName.ToString();
        }
    }
}
