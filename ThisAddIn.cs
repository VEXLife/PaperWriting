using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace PaperWriting
{
    public struct ReferencePreview
    {
        // 提供引用的预览信息
        public enum RefGroup { Formula, Figure, Table } // 引用分类的枚举
        public string Text { get; set; } // 预览中显示的文本
        public Image Image { get; set; } // 预览中显示的图像
        public Word.Bookmark Bookmark { get; set; } // 书签
        public RefGroup Group { get; set; } // 引用分类
    }

    public partial class ThisAddIn
    {
        public About about = new About();
        public CustomTaskPane refTaskPane_pane;
        public RefTaskPane refTaskPane;
        public Properties.Settings Settings = Properties.Settings.Default;
        public SettingsForm settingsForm = new SettingsForm();
        private Word.Selection selection;
        private Word.Document activeDocument;

        #region 插件基本操作
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            refTaskPane = new RefTaskPane();
            refTaskPane_pane = CustomTaskPanes.Add(refTaskPane, "引用");
            refTaskPane_pane.Width = 400;

            Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(UpdateActiveDocument);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Settings.Save();
        }

        private void UpdateActiveDocument()
        {
            selection = Application.Selection;
            try
            {
                activeDocument = Application.ActiveDocument;

                if (Settings.AddTableStyle)
                { // 三线表格样式
                    var styletable = activeDocument.Styles.Add("三线表格", Word.WdStyleType.wdStyleTypeTable);
                    styletable.set_BaseStyle(activeDocument.Styles["普通表格"]);
                    var tableself = styletable.Table;
                    tableself.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tableself.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    tableself.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tableself.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    styletable.Table.Condition(Word.WdConditionCode.wdFirstRow).Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    styletable.Table.Condition(Word.WdConditionCode.wdFirstRow).Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                    activeDocument.UndoClear();
                }
            }
            catch (COMException) { } // 使用try...catch的原因是有可能没有活动的文档
        }
        #endregion

        #region 插入部分
        /// <summary>
        /// 插入带编号的公式。
        /// </summary>
        public void InsertNumberedMath()
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-插入带编号的公式");
            if (Settings.InsertToAnotherParagraph) selection.TypeParagraph();
            AddinUtility.InsertOMath();
            AddinUtility.InsertContent(Settings.Formula, Settings.FormulaStyle);
            Application.UndoRecord.EndCustomRecord();
        }

        /// <summary>
        /// 从文件插入图片。
        /// </summary>
        /// <param name="widthlimit">图片限宽，非正数表示不限宽</param>
        public void InsertFigureFromFile(ref int widthlimit)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-从文件插入带编号的图片");
            OpenFileDialog pickFigure = new OpenFileDialog
            {
                Filter = "所有文件（*.*）|*.*|" +
                "所有图片格式（*.emf;*.wmf;*.jpg;*.jpeg;*.jfif;*.jpe;*.png;*.bmp;*.dib;*.rle;*.gif;*.emz;*.wmz;*.tif;*.tiff;*.svg;*.ico;*.webp）|" +
                "*.emf;*.wmf;*.jpg;*.jpeg;*.jfif;*.jpe;*.png;*.bmp;*.dib;*.rle;*.gif;*.emz;*.wmz;*.tif;*.tiff;*.svg;*.ico;*.webp",
                Title = "插入带编号说明的图片",
                Multiselect = true,
                FilterIndex = 2
            };
            var range = selection.Range;
            if (pickFigure.ShowDialog() == DialogResult.OK)
            {
                bool isFirstOne = true;
                foreach (String filename in pickFigure.FileNames)
                {
                    if (!isFirstOne || Settings.InsertToAnotherParagraph) range.InsertParagraph();
                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    range = AdjustFigure(picture: range.InlineShapes.AddPicture(filename, LinkToFile: false, SaveWithDocument: true),
                        widthlimit: ref widthlimit);
                    isFirstOne = false;
                }
            }
            Application.UndoRecord.EndCustomRecord();
        }

        /// <summary>
        /// 从剪贴板插入图片。
        /// </summary>
        /// <param name="widthlimit">图片限宽，非正数表示不限宽</param>
        public void InsertFigureFromClipboard(ref int widthlimit)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-从剪贴板插入带编号的图片");
            var insertRange = selection.Range;
            selection.Paste();
            insertRange.End = selection.End;

            bool isFirstOne = true;
            foreach (Word.InlineShape pic in insertRange.InlineShapes)
            {
                var range = pic.Range;
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (!isFirstOne || Settings.InsertToAnotherParagraph) range.InsertParagraph();
                AdjustFigure(picture: pic, widthlimit: ref widthlimit);
                isFirstOne = false;
            }
            Application.UndoRecord.EndCustomRecord();
        }

        /// <summary>
        /// 按需求调整图片，包括插入描述。
        /// </summary>
        /// <param name="picture">图片</param>
        /// <param name="widthlimit">图片限宽，非正数表示不限宽</param>
        /// <returns>调整后图片的末尾位置对应的<c>Word.Range</c></returns>
        public Word.Range AdjustFigure(Word.InlineShape picture, ref int widthlimit)
        {
            Word.Range range = picture.Range;

            // 调整大小
            if (widthlimit > 0)
            {
                float ratio = picture.Height / picture.Width;
                picture.Width = widthlimit;
                picture.Height = ratio * widthlimit;
            }
            try
            {
                picture.Range.set_Style(Settings.FigureStyle);
            }
            catch (Exception) { }

            // 根据设置插入描述
            if (Settings.FigurePosition == TargetPosition.Below)
            {
                /* 说明
                 * 想象一下，现在range变量是整个图片，我们要做的是先让光标回到开头。
                 * 按下回车，然后再次让光标回到开头。
                 * 输入图片描述。
                 * 重新把光标移到图片的最后，交给接下来的工作。
                 */
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                range.InsertParagraph();
                range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                AddinUtility.InsertContent(Settings.Figure, Settings.FigureStyle, range);
                range = picture.Range;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
            if (Settings.FigurePosition == TargetPosition.Above)
            {
                /* 说明
                 * 想象一下，现在range变量是整个图片，我们要做的是先让光标前往结尾。
                 * 按下回车，然后再次让光标前往结尾。
                 * 输入图片描述。
                 * 重新把光标移到描述的结尾，交给接下来的工作。
                 */
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraph();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.End = AddinUtility.InsertContent(Settings.Figure, Settings.FigureStyle, range).End;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
            return range;
        }
        #endregion

        #region 引用管理
        /// <summary>
        /// 获取引用的预览。
        /// </summary>
        /// <param name="imgWidth">引用预览图片的宽度</param>
        /// <param name="imgHeight">引用预览图片的高度</param>
        /// <returns>一个<c>List</c>，其中是若干预览信息</returns>
        public List<ReferencePreview> GetReferencePreviews(int imgWidth = 400, int imgHeight = 100)
        {
            List<ReferencePreview> previews = new List<ReferencePreview>();
            if (activeDocument != null)
                foreach (Word.Bookmark bookmark in activeDocument.Bookmarks)
                {
                    if (!bookmark.Name.StartsWith(Settings.BookmarkPrefix)) continue; // 只处理前缀正确的书签
                    try
                    {
                        if (bookmark.Range.OMaths.Count > 0)
                        { // 公式
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
                            ReferencePreview preview = new ReferencePreview
                            {
                                Text = bookmark.Range.Text,
                                Image = bmp,
                                Bookmark = bookmark,
                                Group = ReferencePreview.RefGroup.Formula
                            };
                            previews.Add(preview);
                        }
                        else if ((Settings.FigurePosition == TargetPosition.Below
                        && bookmark.Range.Paragraphs[1].Next() != null
                        && bookmark.Range.Paragraphs[1].Next().Range.InlineShapes.Count > 0) ||
                        (Settings.FigurePosition == TargetPosition.Above
                        && bookmark.Range.Paragraphs[1].Previous() != null
                        && bookmark.Range.Paragraphs[1].Previous().Range.InlineShapes.Count > 0))
                        { // 图片
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
                            ReferencePreview preview = new ReferencePreview
                            {
                                Text = bookmark.Range.Paragraphs[1].Range.Text,
                                Image = bmp,
                                Bookmark = bookmark,
                                Group = ReferencePreview.RefGroup.Figure
                            };
                            previews.Add(preview);
                        }
                        else if ((Settings.TablePosition == TargetPosition.Below
                        && bookmark.Range.Paragraphs[1].Next() != null
                        && bookmark.Range.Paragraphs[1].Next().Range.Tables.Count > 0) ||
                        (Settings.TablePosition == TargetPosition.Above
                        && bookmark.Range.Paragraphs[1].Previous() != null
                        && bookmark.Range.Paragraphs[1].Previous().Range.Tables.Count > 0))
                        { // 表格
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
                            ReferencePreview preview = new ReferencePreview
                            {
                                Text = bookmark.Range.Paragraphs[1].Range.Text,
                                Image = bmp,
                                Bookmark = bookmark,
                                Group = ReferencePreview.RefGroup.Table
                            };
                            previews.Add(preview);
                        }
                    }
                    catch (COMException) { }
                }
            return previews;
        }

        /// <summary>
        /// 添加引用。
        /// </summary>
        /// <param name="bookmarkName">所引的书签名称</param>
        /// <param name="hyperref">是否链接过去</param>
        public void AddRef(string bookmarkName, bool hyperref = true)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-引用内容");
            var selection = Application.Selection;
            selection.InsertCrossReference("书签", Word.WdReferenceKind.wdContentText, bookmarkName, hyperref);
            Application.UndoRecord.EndCustomRecord();
        }

        /// <summary>
        /// 删除书签。
        /// </summary>
        /// <param name="bookmarkName">书签名称</param>
        public void RemoveBookmark(string bookmarkName)
        {
            Application.UndoRecord.StartCustomRecord("论文辅助-删除可引用的项");
            Word.Document document = Application.ActiveDocument;
            document.Bookmarks[bookmarkName].Delete();
            Application.UndoRecord.EndCustomRecord();
        }
        #endregion

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
