﻿using System;
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
        public enum RefGroup { Formula, Figure, Table }
        public string Text { get; set; }
        public Image Image { get; set; }
        public Word.Bookmark Bookmark { get; set; }
        public RefGroup Group { get; set; }
    }

    public partial class ThisAddIn
    {
        public About about = new About();
        public CustomTaskPane refTaskPane_pane;
        public RefTaskPane refTaskPane;
        public Properties.Settings Settings = Properties.Settings.Default;
        public SettingsForm settingsForm = new SettingsForm();
        private Word.Selection selection;
        private Word.Document document;

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

            selection = Application.Selection;
            this.document = document;
        }

        public void InsertNumberedMath()
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-插入带编号的公式");
            selection.TypeParagraph();
            AddinUtility.InsertOMath();
            AddinUtility.InsertContent(Settings.Formula, Settings.FormulaStyle);
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
            var range = selection.Range;
            if (pickFigure.ShowDialog() == DialogResult.OK)
            {
                foreach (String filename in pickFigure.FileNames)
                {
                    range.InsertParagraph();
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    range = AdjustFigure(picture: range.InlineShapes.AddPicture(filename, LinkToFile: false, SaveWithDocument: true),
                        widthlimit: ref widthlimit);
                }
            }
            Application.UndoRecord.EndCustomRecord();
        }

        public void InsertFigureFromClipboard(ref int widthlimit)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-从剪贴板插入带编号的图片");
            var insertRange = selection.Range;
            selection.Paste();
            insertRange.End = selection.End;

            foreach (Word.InlineShape pic in insertRange.InlineShapes)
            {
                var range = pic.Range;
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.InsertParagraph();
                AdjustFigure(picture: pic, widthlimit: ref widthlimit);
            }
            Application.UndoRecord.EndCustomRecord();
        }

        public Word.Range AdjustFigure(Word.InlineShape picture, ref int widthlimit)
        {
            Word.Range range = picture.Range;
            if (widthlimit > 0)
            {
                float ratio = picture.Height / picture.Width;
                picture.Width = widthlimit;
                picture.Height = ratio * widthlimit;
            }
            picture.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            if (Settings.FigurePosition == TargetPosition.Below)
            {
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.InsertParagraph();
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                AddinUtility.InsertContent(Settings.Figure, Settings.FigureStyle, range);
                range = picture.Range;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
            if (Settings.FigurePosition == TargetPosition.Above)
            {
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraph();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.End = AddinUtility.InsertContent(Settings.Figure, Settings.FigureStyle, range).End;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
            return range;
        }

        public List<QuotePreview> GetQuotePreviews(int imgWidth = 400, int imgHeight = 100)
        {
            List<QuotePreview> previews = new List<QuotePreview>();
            if (document != null)
                foreach (Word.Bookmark bookmark in document.Bookmarks)
                {
                    if (!bookmark.Name.StartsWith(Settings.BookmarkPrefix)) continue;
                    if (bookmark.Range.OMaths.Count > 0)
                    {
                        QuotePreview preview = new QuotePreview();
                        preview.Text = bookmark.Range.Text;
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
                        preview.Group = QuotePreview.RefGroup.Formula;
                        previews.Add(preview);
                    }
                    else if ((Settings.FigurePosition == TargetPosition.Below
                    && bookmark.Range.Paragraphs[1].Next() != null
                    && bookmark.Range.Paragraphs[1].Next().Range.InlineShapes.Count > 0) ||
                    (Settings.FigurePosition == TargetPosition.Above
                    && bookmark.Range.Paragraphs[1].Previous() != null
                    && bookmark.Range.Paragraphs[1].Previous().Range.InlineShapes.Count > 0))
                    {
                        QuotePreview preview = new QuotePreview();
                        preview.Text = bookmark.Range.Paragraphs[1].Range.Text;
                        Image enhImage = Image.FromStream(
                            new System.IO.MemoryStream(
                                (byte[])(Settings.FigurePosition == TargetPosition.Below ?
                                bookmark.Range.Paragraphs[1].Next().Range.InlineShapes[1].Range.EnhMetaFileBits :
                                bookmark.Range.Paragraphs[1].Previous().Range.InlineShapes[1].Range.EnhMetaFileBits)
                            )
                        );
                        Bitmap bmp = new Bitmap(imgWidth, imgHeight);
                        Graphics pen = Graphics.FromImage(bmp);
                        pen.DrawImage(enhImage, 0, 0);
                        preview.Image = bmp;
                        preview.Bookmark = bookmark;
                        previews.Add(preview);
                    }
                    else if ((Settings.TablePosition == TargetPosition.Below
                    && bookmark.Range.Paragraphs[1].Next() != null
                    && bookmark.Range.Paragraphs[1].Next().Range.Tables.Count > 0) ||
                    (Settings.TablePosition == TargetPosition.Above
                    && bookmark.Range.Paragraphs[1].Previous() != null
                    && bookmark.Range.Paragraphs[1].Previous().Range.Tables.Count > 0))
                    {
                        QuotePreview preview = new QuotePreview();
                        preview.Text = bookmark.Range.Paragraphs[1].Range.Text;
                        Image enhImage = Image.FromStream(
                            new System.IO.MemoryStream(
                                (byte[])(Settings.TablePosition == TargetPosition.Below ?
                                bookmark.Range.Paragraphs[1].Next().Range.Tables[1].Range.EnhMetaFileBits :
                                bookmark.Range.Paragraphs[1].Previous().Range.Tables[1].Range.EnhMetaFileBits)
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
            return previews;
        }

        public void AddRef(string bookmarkName, bool hyperref = true)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-引用内容");
            var selection = Application.Selection;
            selection.InsertCrossReference("书签", Word.WdReferenceKind.wdContentText, bookmarkName, hyperref);
            Application.UndoRecord.EndCustomRecord();
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