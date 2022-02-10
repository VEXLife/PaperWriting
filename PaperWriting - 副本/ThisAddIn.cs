using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Office.Tools;

namespace PaperWriting
{
    public struct QuotePreview
    {
        public string Text { get; set; }
        public Image Image { get; set; }
        public Word.Bookmark Bookmark { get; set; }
    }

    public partial class ThisAddIn
    {
        static int[] bookmarks = new int[3];
        static string[] descriptions = new string[] { "", "Fig.", "Table " };
        static string[] prefixes = new string[] { "Equation_", "Figure_", "Table_" };
        static string[] SEQs = new string[] { "公式", "图片", "表格" };
        public About about = new About();
        public CustomTaskPane refTaskPane_pane;
        public RefTaskPane refTaskPane;
        public Properties.Settings Settings = new Properties.Settings();
        public SettingsForm settingsForm = new SettingsForm();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(UpdateBookmarkIndex);

            refTaskPane = new RefTaskPane();
            refTaskPane_pane = CustomTaskPanes.Add(refTaskPane, "引用");
            refTaskPane_pane.Width = 400;

            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(UpdateBookmarkIndex);
        }

        private void UpdateBookmarkIndex(Word.Document document)
        {
            while (document.Bookmarks.Exists(prefixes[0] + bookmarks[0].ToString()))
            {
                bookmarks[0]++;
            }
            while (document.Bookmarks.Exists(prefixes[1] + bookmarks[1].ToString()))
            {
                bookmarks[1]++;
            }
            while (document.Bookmarks.Exists(prefixes[2] + bookmarks[2].ToString()))
            {
                bookmarks[2]++;
            }

            var styles = document.Styles.Add("Pictures and Figures", Word.WdStyleType.wdStyleTypeParagraph);
            styles.Font.Size = 12;
            styles.Font.Name = "Times New Roman";
            styles.Font.Italic = 1;
            styles.Font.Bold = 1;
            styles.Font.Color = Word.WdColor.wdColorBlack;
            styles.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            styles.set_NextParagraphStyle(document.Styles["正文"]);

            var styletable = document.Styles.Add("三线表格", Word.WdStyleType.wdStyleTypeTable);
            styletable.set_BaseStyle(document.Styles["普通表格"]);
            var tableself = styletable.Table;
            tableself.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            tableself.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
            tableself.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            tableself.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
            styletable.Table.Condition(Word.WdConditionCode.wdFirstRow).Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            styletable.Table.Condition(Word.WdConditionCode.wdFirstRow).Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
            document.UndoClear();
        }

        public void InsertOMath()
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-插入带编号的公式");
            var selection = Application.Selection;
            Word.Document document = Application.ActiveDocument;
            selection.TypeParagraph();
            document.OMaths.Add(selection.Range);
            selection.TypeText("#(");
            Word.Range result = document.Fields.Add(selection.Range, Word.WdFieldType.wdFieldEmpty, "SEQ " + SEQs[0], false).Result;
            selection.TypeText(")");
            selection.MoveLeft(Unit: Word.WdUnits.wdWord, 3);
            result.Start--;
            result.End++;
            document.Bookmarks.Add(prefixes[0] + bookmarks[0].ToString(), result);
            bookmarks[0]++;
            refTaskPane.RefreshContent();
            Application.UndoRecord.EndCustomRecord();
        }

        public void InsertFigureFromFile(ref int widthlimit)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-从文件插入带编号的图片");
            OpenFileDialog pickFigure = new OpenFileDialog();
            pickFigure.Filter = "所有文件（*.*）|*.*|" +
                "所有图片格式（*.emf;*.wmf;*.jpg;*.jpeg;*.jfif;*.jpe;*.png;*.bmp;*.dib;*.rle;*.gif;*.emz;*.wmz;*.tif;*.tiff;*.svg;*.ico;*.webp）|" +
                "*.emf;*.wmf;*.jpg;*.jpeg;*.jfif;*.jpe;*.png;*.bmp;*.dib;*.rle;*.gif;*.emz;*.wmz;*.tif;*.tiff;*.svg;*.ico;*.webp";
            pickFigure.Title = "插入带编号说明的图片";
            pickFigure.Multiselect = true;
            pickFigure.FilterIndex = 2;
            if (pickFigure.ShowDialog() == DialogResult.OK)
            {
                foreach (String filename in pickFigure.FileNames)
                {
                    InsertFigure(Filename: filename, widthlimit: ref widthlimit);
                }
            }
            Application.UndoRecord.EndCustomRecord();
        }

        public void InsertFigureFromClipboard(ref int widthlimit)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-从剪贴板插入带编号的图片");
            var selection = Application.Selection;
            selection.set_Style(Application.ActiveDocument.Styles["Pictures and Figures"]);
            var insertedRange = selection.Range;
            selection.Paste();
            insertedRange.End = selection.End;

            var shapes = insertedRange.InlineShapes;
            foreach (Word.InlineShape pic in shapes)
            {
                if (widthlimit > 0)
                {
                    float ratio = pic.Height / pic.Width;
                    pic.Width = widthlimit;
                    pic.Height = ratio * widthlimit;
                }
                var range = pic.Range;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphBefore();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                AddLabel(1, range).InsertParagraph();
            }
            selection.Delete();
            Application.UndoRecord.EndCustomRecord();
        }

        public void InsertFigure(String Filename, ref int widthlimit)
        {
            Word.Document document = Application.ActiveDocument;
            var selection = Application.Selection;
            selection.TypeParagraph();
            selection.set_Style(document.Styles["Pictures and Figures"]);
            var pic = selection.InlineShapes.AddPicture(Filename, LinkToFile: false, SaveWithDocument: true);
            if (widthlimit > 0)
            {
                float ratio = pic.Height / pic.Width;
                pic.Width = widthlimit;
                pic.Height = ratio * widthlimit;
            }
            selection.TypeParagraph();
            AddLabel(1, selection.Range);
        }

        public Word.Range AddLabel(int type, Word.Range range)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-插入描述");
            Word.Document document = Application.ActiveDocument;
            range.set_Style(document.Styles["Pictures and Figures"]);
            range.InsertBefore(descriptions[type]);
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Range result = document.Fields.Add(range, Word.WdFieldType.wdFieldEmpty, "SEQ " + SEQs[type], false).Result;
            document.Bookmarks.Add(prefixes[type] + bookmarks[type].ToString(), result);
            result.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            result.Move(Count: 1);
            result.InsertBefore("  ");
            bookmarks[type]++;
            refTaskPane.RefreshContent();
            Application.Selection.SetRange(result.End, result.End);
            Application.UndoRecord.EndCustomRecord();
            return document.Range(result.End, result.End);
        }

        public List<QuotePreview> GetQuotePreviews(int imgWidth = 400, int imgHeight = 100)
        {
            Word.Document document = Application.ActiveDocument;
            List<QuotePreview> previews = new List<QuotePreview>();
            foreach (Word.Bookmark bookmark in document.Bookmarks)
            {
                try
                {
                    if (bookmark.Name.StartsWith(prefixes[0]))
                    {
                        QuotePreview preview = new QuotePreview();
                        preview.Text = "公式" + bookmark.Range.Text;
                        Image enhImage = Image.FromStream(
                            new System.IO.MemoryStream(
                                (byte[])bookmark.Range.Paragraphs[1].Range.OMaths[1].Range.EnhMetaFileBits
                            )
                        );
                        Bitmap bmp = new Bitmap(imgWidth, imgHeight);
                        Graphics pen = Graphics.FromImage(bmp);
                        pen.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        int width = enhImage.Width > (imgWidth / imgHeight * enhImage.Height) ? (imgWidth / imgHeight * enhImage.Height) : enhImage.Width;
                        var rect = new Rectangle(0, 0, imgWidth, imgHeight);
                        pen.DrawImage(enhImage, rect, enhImage.Width / 2 - width / 2, 0, width, enhImage.Height, GraphicsUnit.Pixel);
                        preview.Image = bmp;
                        preview.Bookmark = bookmark;
                        previews.Add(preview);
                    }
                    else if (bookmark.Name.StartsWith(prefixes[1]))
                    {
                        QuotePreview preview = new QuotePreview();
                        preview.Text = bookmark.Range.Paragraphs[1].Range.Text;
                        Image enhImage = Image.FromStream(
                            new System.IO.MemoryStream(
                                (byte[])bookmark.Range.Paragraphs[1].Previous().Range.InlineShapes[1].Range.EnhMetaFileBits
                            )
                        );
                        Bitmap bmp = new Bitmap(imgWidth, imgHeight);
                        Graphics pen = Graphics.FromImage(bmp);
                        pen.DrawImage(enhImage, 0, 0);
                        preview.Image = bmp;
                        preview.Bookmark = bookmark;
                        previews.Add(preview);
                    }
                    else if (bookmark.Name.StartsWith(prefixes[2]))
                    {
                        QuotePreview preview = new QuotePreview();
                        preview.Text = bookmark.Range.Paragraphs[1].Range.Text;
                        Image enhImage = Image.FromStream(
                            new System.IO.MemoryStream(
                                (byte[])bookmark.Range.Paragraphs[1].Next().Range.Tables[1].Range.EnhMetaFileBits
                            )
                        );
                        Bitmap bmp = new Bitmap(imgWidth, imgHeight);
                        Graphics pen = Graphics.FromImage(bmp);
                        pen.DrawImage(enhImage, 0, 0);
                        preview.Image = bmp;
                        preview.Bookmark = bookmark;
                        previews.Add(preview);
                    }
                }
                catch (Exception) { }
            }
            return previews;
        }

        public void AddRef(string bookmarkName, bool hyperref = true)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-引用内容");
            var selection = Application.Selection;
            selection.InsertCrossReference("书签", Word.WdReferenceKind.wdContentText, bookmarkName, hyperref);
            Application.UndoRecord.EndCustomRecord();
        }

        public string CatagotizeBookmark(string bookmarkName)
        {
            for (int i = 0; i < 3; i++)
            {
                if (bookmarkName.StartsWith(prefixes[i]))
                    return SEQs[i];
            }
            return "未知类别";
        }

        public string CatagorizeBookmark(Word.Bookmark bookmarkName)
        {
            return CatagotizeBookmark(bookmarkName.Name);
        }

        public void RemoveBookmark(string bookmarkName)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-删除可引用的项");
            Word.Document document = Application.ActiveDocument;
            document.Bookmarks[bookmarkName].Delete();
            Application.UndoRecord.EndCustomRecord();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Settings.Save();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
