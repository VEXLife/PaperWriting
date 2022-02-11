using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PaperWriting
{
    public partial class RefTaskPane : UserControl
    {
        private Properties.Settings Settings=Properties.Settings.Default;
     
        public RefTaskPane()
        {
            InitializeComponent();
        }

        public void RefreshContent()
        {
            try
            {
                hyperref.Checked = Settings.HyperRef;
                var previews = Globals.ThisAddIn.GetQuotePreviews();
                imageList.Images.Clear();
                QuotableContents.Clear();
                foreach (var preview in previews)
                {
                    imageList.Images.Add(preview.Bookmark.Name, preview.Image);
                    ListViewItem item = new ListViewItem();
                    item.Text = preview.Text;
                    item.ImageKey = preview.Bookmark.Name;
                    item.Group = QuotableContents.Groups[Globals.ThisAddIn.CatagorizeBookmark(preview.Bookmark)];
                    item.Tag = preview.Bookmark.Name;
                    QuotableContents.Items.Add(item);
                }
            }
            catch (Exception) { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RefreshContent();
        }

        private void RefTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            if (Visible == true)
                RefreshContent();
        }

        private void insert_Click(object sender, EventArgs e)
        {
            insert_selected();
        }

        private void insert_selected()
        {
            foreach (var selectedItem in QuotableContents.SelectedItems)
            {
                Globals.ThisAddIn.AddRef((string)((ListViewItem)selectedItem).Tag, hyperref.Checked);
            }
        }

        private void delete_selected()
        {
            foreach (var selectedItem in QuotableContents.SelectedItems)
            {
                Globals.ThisAddIn.RemoveBookmark((string)((ListViewItem)selectedItem).Tag);
                QuotableContents.Items.Remove((ListViewItem)selectedItem);
            }
        }

        private void delete_Click(object sender, EventArgs e)
        {
            delete_selected();
        }

        private void QuotableContents_DoubleClick(object sender, EventArgs e)
        {
            insert_selected();
        }

        private void 引用选中项ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            insert_selected();
        }

        private void 删除选中项ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            delete_selected();
        }

        private void 刷新ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RefreshContent();
        }

        private void onKeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                insert_selected();
            }
        }

        private void hyperref_CheckedChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Settings.HyperRef = hyperref.Checked;
        }

        private void button1_MouseDown(object sender, MouseEventArgs e)
        {
            ((Button)sender).FlatAppearance.BorderColor = Color.FromArgb(62, 109, 181);
        }

        private void button1_MouseUp(object sender, MouseEventArgs e)
        {
            ((Button)sender).FlatAppearance.BorderColor = Color.FromArgb(163, 189, 227);
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            ((Button)sender).FlatAppearance.BorderColor = Color.FromArgb(163, 189, 227);
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            ((Button)sender).FlatAppearance.BorderColor = Color.FromArgb(171,171,171);
        }
    }
}
