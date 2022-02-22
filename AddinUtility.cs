using Microsoft.Office.Interop.Word;
using System;
using System.Text;
using System.Text.RegularExpressions;

namespace PaperWriting
{
    public class AddinUtility
    {
        private const string bracketRegex = @"((?<!\\)\{.*?(?<!\\)\}|(?<!\\)[\[\]]|(?<!\\)#)"; // 用于匹配可识别的功能字符的正则表达式

        private static Selection selection { get { return Globals.ThisAddIn.Application.Selection; } }
        private static Document document { get { return Globals.ThisAddIn.Application.ActiveDocument; } }
        private static UndoRecord undoRecord { get { return Globals.ThisAddIn.Application.UndoRecord; } }

        /// <summary>
        /// 插入公式。参见<seealso cref="InsertContent(string, string, Range)"/>
        /// </summary>
        /// <param name="range">插入的位置</param>
        /// <returns>插入后公式的范围</returns>
        public static Range InsertOMath(Range range = null)
        {
            if (range == null) range = selection.Range;
            return range.OMaths.Add(range);
        }

        /// <summary>
        /// 插入域代码。参见<seealso cref="InsertContent(string, string, Range)"/>
        /// </summary>
        /// <param name="code">域代码</param>
        /// <param name="range">插入的位置</param>
        /// <returns>插入后域的范围</returns>
        public static Range InsertField(string code, Range range = null)
        {
            if (range == null) range = selection.Range;
            return range.Fields.Add(range, WdFieldType.wdFieldEmpty, code, false).Result;
        }

        /// <summary>
        /// 插入插件支持的内容。
        /// </summary>
        /// <param name="content">内容，允许包含可识别的字符</param>
        /// <param name="style">要设置的样式</param>
        /// <param name="range">插入的位置</param>
        /// <returns>插入后最后的字符位置</returns>
        public static Range InsertContent(string content, string style = null, Range range = null)
        {
            undoRecord.StartCustomRecord("论文辅助-插入文本");

            if (range == null) range = selection.Range;
            try
            {
                range.set_Style(style);
            }
            catch (System.Runtime.InteropServices.COMException) { }
            Range bookmarkRange = document.Range();

            var contentTextArray = Regex.Split(content, bracketRegex);
            foreach (string contentText in contentTextArray)
            {
                if (contentText == "") continue;
                if (contentText.StartsWith("{") && contentText.EndsWith("}"))
                {
                    /* 说明
                     * 存在一个问题：Word插入域代码后返回的范围是考虑域代码的长度的，而切换域代码后这个范围就不对了。
                     * 目前的解决方案是把原本的范围右移返回的范围长度，得到插入域后的位置。
                     */
                    var fieldRange = InsertField(contentText.Substring(1, contentText.Length - 2), range);
                    range.Move(Count: fieldRange.End - fieldRange.Start);
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
                    selection.SetRange(range.Start, range.Start);
                    continue;
                }
                range.InsertAfter(contentText.Replace(@"\#", "#").Replace(@"\{", "{").Replace(@"\}", "}").Replace(@"\[", "[").Replace(@"\]", "]")); // 插入纯文本内容，还原转义字符
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
            }
            document.Bookmarks.Add(GenerateBookmarkName(), bookmarkRange);
            Globals.ThisAddIn.refTaskPane.RefreshContent();
            undoRecord.EndCustomRecord();
            return range;
        }

        /// <summary>
        /// 生成一个书签名称。代码参考了互联网上的一篇博客，地址：https://www.cnblogs.com/binger333/p/4693757.html
        /// </summary>
        /// <returns>书签名称</returns>
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
                bookmarkName = new StringBuilder(Properties.Settings.Default.BookmarkPrefix); // 带上前缀便于标识
                for (int i = 0; i < 12; i++)
                {
                    bookmarkName.Append(constant[random.Next(constant.Length)]);
                }
            } while (document.Bookmarks.Exists(bookmarkName.ToString()));
            return bookmarkName.ToString();
        }
    }
}
