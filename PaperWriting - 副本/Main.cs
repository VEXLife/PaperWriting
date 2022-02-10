using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace PaperWriting
{
    public partial class Main
    {

        private void inmaths_Click(object sender, RibbonControlEventArgs e) => Globals.ThisAddIn.InsertOMath();

        private void insert_figs_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            int widthlimit = -1;
            try
            {
                if (this.widthlimit.Text != "")
                {
                    widthlimit = int.Parse(this.widthlimit.Text);
                }
            }
            catch (FormatException) { }
            if (((RibbonControl)sender).Id == "from_file" || ((RibbonControl)sender).Id == "insert_figs")
            {
                Globals.ThisAddIn.InsertFigureFromFile(ref widthlimit);
            }
            else if (((RibbonControl)sender).Id == "from_clipboard")
            {
                Globals.ThisAddIn.InsertFigureFromClipboard(ref widthlimit);
            }
        }

        private void insert_label_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            if (((RibbonControl)sender).Id == "figlabel")
            {
                Globals.ThisAddIn.AddLabel(1, Globals.ThisAddIn.Application.Selection.Range);
            }
            else if (((RibbonControl)sender).Id == "tablelabel" || ((RibbonControl)sender).Id == "insert_label")
            {
                Globals.ThisAddIn.AddLabel(2, Globals.ThisAddIn.Application.Selection.Range);
            }
        }

        private void addQuote_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            addQuote.Items.Clear();
            foreach (var quoteItem in Globals.ThisAddIn.GetQuotePreviews())
            {
                RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                item.Label = quoteItem.Text;
                item.Image = quoteItem.Image;
                item.Tag = quoteItem.Bookmark.Name;
                addQuote.Items.Add(item);
            }
        }

        private void addQuote_Click(object sender, RibbonControlEventArgs e) => Globals.ThisAddIn.AddRef((string)addQuote.SelectedItem.Tag);

        private void widthlimit_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (widthlimit.Text != "" && !int.TryParse(widthlimit.Text, out _))
                widthlimit.Text = "";
            Globals.ThisAddIn.Settings.WidthLimit = widthlimit.Text;
        }

        private void quotes_DialogLauncherClick(object sender, RibbonControlEventArgs e) => Globals.ThisAddIn.refTaskPane_pane.Visible = !Globals.ThisAddIn.refTaskPane_pane.Visible;

        private void addQuote_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            if (e.Control.Id == "show_refTaskPane")
                Globals.ThisAddIn.refTaskPane_pane.Visible = !Globals.ThisAddIn.refTaskPane_pane.Visible;
        }

        private void Main_Load(object sender, RibbonUIEventArgs e)
        {
            widthlimit.Text = Globals.ThisAddIn.Settings.WidthLimit;
        }

        private void Insert_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.settingsForm.IsDisposed) Globals.ThisAddIn.settingsForm = new SettingsForm();
            Globals.ThisAddIn.settingsForm.Show();
            Globals.ThisAddIn.settingsForm.Focus();
        }

        private void headers_ButtonClick(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
