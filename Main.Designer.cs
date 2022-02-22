namespace PaperWriting
{
    partial class Main : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Main()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl2 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Insert = this.Factory.CreateRibbonGroup();
            this.headers = this.Factory.CreateRibbonGallery();
            this.h1 = this.Factory.CreateRibbonButton();
            this.h2 = this.Factory.CreateRibbonButton();
            this.h3 = this.Factory.CreateRibbonButton();
            this.insert_figs = this.Factory.CreateRibbonSplitButton();
            this.from_file = this.Factory.CreateRibbonButton();
            this.from_clipboard = this.Factory.CreateRibbonButton();
            this.widthlimit = this.Factory.CreateRibbonEditBox();
            this.inmaths = this.Factory.CreateRibbonButton();
            this.insert_label = this.Factory.CreateRibbonSplitButton();
            this.tablelabel = this.Factory.CreateRibbonButton();
            this.figlabel = this.Factory.CreateRibbonButton();
            this.makequotable = this.Factory.CreateRibbonButton();
            this.quotes = this.Factory.CreateRibbonGroup();
            this.addQuote = this.Factory.CreateRibbonGallery();
            this.show_refTaskPane = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Insert.SuspendLayout();
            this.quotes.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Insert);
            this.tab1.Groups.Add(this.quotes);
            this.tab1.Label = "论文辅助";
            this.tab1.Name = "tab1";
            // 
            // Insert
            // 
            ribbonDialogLauncherImpl1.ScreenTip = "打开加载项设置";
            ribbonDialogLauncherImpl1.SuperTip = "调整加载项快捷插入内容时的行为。加载项设置中还包括一个“关于”页面。";
            this.Insert.DialogLauncher = ribbonDialogLauncherImpl1;
            this.Insert.Items.Add(this.headers);
            this.Insert.Items.Add(this.insert_figs);
            this.Insert.Items.Add(this.widthlimit);
            this.Insert.Items.Add(this.inmaths);
            this.Insert.Items.Add(this.insert_label);
            this.Insert.Items.Add(this.makequotable);
            this.Insert.Label = "论文插入";
            this.Insert.Name = "Insert";
            this.Insert.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Insert_DialogLauncherClick);
            // 
            // headers
            // 
            this.headers.Buttons.Add(this.h1);
            this.headers.Buttons.Add(this.h2);
            this.headers.Buttons.Add(this.h3);
            this.headers.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.headers.Image = global::PaperWriting.Properties.Resources.icons8_header_1_40;
            this.headers.Label = "添加标题";
            this.headers.Name = "headers";
            this.headers.ScreenTip = "插入一个标题";
            this.headers.ShowImage = true;
            this.headers.SuperTip = "可提供自动输入所设置的标题部分内容的功能。";
            this.headers.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.headers_ButtonClick);
            // 
            // h1
            // 
            this.h1.Image = global::PaperWriting.Properties.Resources.icons8_header_1_40;
            this.h1.Label = "一级标题";
            this.h1.Name = "h1";
            this.h1.ShowImage = true;
            this.h1.SuperTip = "一般是最主要的标题层次。";
            // 
            // h2
            // 
            this.h2.Image = global::PaperWriting.Properties.Resources.icons8_header_2_40;
            this.h2.Label = "二级标题";
            this.h2.Name = "h2";
            this.h2.ShowImage = true;
            this.h2.SuperTip = "一般是次级标题，如小节等。";
            // 
            // h3
            // 
            this.h3.Image = global::PaperWriting.Properties.Resources.icons8_header_3_40;
            this.h3.Label = "三级标题";
            this.h3.Name = "h3";
            this.h3.ShowImage = true;
            this.h3.SuperTip = "一般是最小的标题层次，如方法或例子等。";
            // 
            // insert_figs
            // 
            this.insert_figs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.insert_figs.Image = global::PaperWriting.Properties.Resources.icons8_add_image_48;
            this.insert_figs.Items.Add(this.from_file);
            this.insert_figs.Items.Add(this.from_clipboard);
            this.insert_figs.Label = "带编号图片";
            this.insert_figs.Name = "insert_figs";
            this.insert_figs.ScreenTip = "插入带编号的图片";
            this.insert_figs.SuperTip = "插入图片后将自动添加图片描述，并应用所设置的图片样式。直接点击为“来自文件”。";
            this.insert_figs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insert_figs_ButtonClick);
            // 
            // from_file
            // 
            this.from_file.Image = global::PaperWriting.Properties.Resources.icons8_图像文件_48;
            this.from_file.Label = "来自文件";
            this.from_file.Name = "from_file";
            this.from_file.ScreenTip = "从文件插入图片";
            this.from_file.ShowImage = true;
            this.from_file.SuperTip = "点击后需选取图片文件，允许多选。";
            this.from_file.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insert_figs_ButtonClick);
            // 
            // from_clipboard
            // 
            this.from_clipboard.Image = global::PaperWriting.Properties.Resources.icons8_粘贴_48;
            this.from_clipboard.Label = "来自剪贴板";
            this.from_clipboard.Name = "from_clipboard";
            this.from_clipboard.ScreenTip = "从剪贴板插入图片";
            this.from_clipboard.ShowImage = true;
            this.from_clipboard.SuperTip = "您需要先复制一张图片，暂不支持多张。";
            this.from_clipboard.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insert_figs_ButtonClick);
            // 
            // widthlimit
            // 
            this.widthlimit.Image = global::PaperWriting.Properties.Resources.icons8_调整水平_48;
            this.widthlimit.Label = "限宽：";
            this.widthlimit.Name = "widthlimit";
            this.widthlimit.ScreenTip = "限制插入的图片宽度";
            this.widthlimit.ShowImage = true;
            this.widthlimit.SizeString = "0000";
            this.widthlimit.SuperTip = "不填代表不限，填了高度按等比例调整。";
            this.widthlimit.Text = null;
            this.widthlimit.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.widthlimit_TextChanged);
            // 
            // inmaths
            // 
            this.inmaths.Image = global::PaperWriting.Properties.Resources.icons8_pi_48;
            this.inmaths.Label = "带编号公式";
            this.inmaths.Name = "inmaths";
            this.inmaths.ScreenTip = "插入带编号的公式";
            this.inmaths.ShowImage = true;
            this.inmaths.SuperTip = "可提供自动输入所设置的公式部分内容的功能。";
            this.inmaths.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.inmaths_Click);
            // 
            // insert_label
            // 
            this.insert_label.Image = global::PaperWriting.Properties.Resources.text;
            this.insert_label.Items.Add(this.tablelabel);
            this.insert_label.Items.Add(this.figlabel);
            this.insert_label.Label = "添加描述";
            this.insert_label.Name = "insert_label";
            this.insert_label.ScreenTip = "添加描述文本";
            this.insert_label.SuperTip = "描述文本将以所设置的样式应用并呈现。直接点击为表格描述。";
            this.insert_label.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insert_label_ButtonClick);
            // 
            // tablelabel
            // 
            this.tablelabel.Image = global::PaperWriting.Properties.Resources.icons8_插入表格_48;
            this.tablelabel.Label = "表格描述";
            this.tablelabel.Name = "tablelabel";
            this.tablelabel.ScreenTip = "添加表格描述";
            this.tablelabel.ShowImage = true;
            this.tablelabel.SuperTip = "您需要在目标表格上方或下方使用，取决于您设置的结果。";
            this.tablelabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insert_label_ButtonClick);
            // 
            // figlabel
            // 
            this.figlabel.Image = global::PaperWriting.Properties.Resources.icons8_图像_48;
            this.figlabel.Label = "图片描述";
            this.figlabel.Name = "figlabel";
            this.figlabel.ScreenTip = "添加图片描述";
            this.figlabel.ShowImage = true;
            this.figlabel.SuperTip = "您需要在目标图片上方或下方使用，取决于您设置的结果。";
            this.figlabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insert_label_ButtonClick);
            // 
            // makequotable
            // 
            this.makequotable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.makequotable.Image = global::PaperWriting.Properties.Resources.icons8_bookmark_48;
            this.makequotable.Label = "标记为可引用";
            this.makequotable.Name = "makequotable";
            this.makequotable.ScreenTip = "将选中内容标记为将在引用窗格中出现的内容";
            this.makequotable.ShowImage = true;
            this.makequotable.SuperTip = "如果您遇到了意外的错误，或是无意间删除了不希望删除的引用标记，抑或是有加载项尚不支持自动插入的可引用内容，均可使用此功能。";
            this.makequotable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.makequotable_Click);
            // 
            // quotes
            // 
            ribbonDialogLauncherImpl2.Image = global::PaperWriting.Properties.Resources.icons8_get_quote_48;
            ribbonDialogLauncherImpl2.ScreenTip = "打开引用窗格";
            ribbonDialogLauncherImpl2.SuperTip = "该任务窗格可更进一步方便您大量引用的操作。";
            this.quotes.DialogLauncher = ribbonDialogLauncherImpl2;
            this.quotes.Items.Add(this.addQuote);
            this.quotes.Label = "引用";
            this.quotes.Name = "quotes";
            this.quotes.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.quotes_DialogLauncherClick);
            // 
            // addQuote
            // 
            this.addQuote.Buttons.Add(this.show_refTaskPane);
            this.addQuote.ColumnCount = 1;
            this.addQuote.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.addQuote.Image = global::PaperWriting.Properties.Resources.icons8_quote_48;
            this.addQuote.Label = "添加引用";
            this.addQuote.Name = "addQuote";
            this.addQuote.ScreenTip = "添加指向您创建的带编号内容的引用";
            this.addQuote.ShowImage = true;
            this.addQuote.SuperTip = "编号内容可以是由本加载项创建的，或是您自己手动添加的书签，书签名必须以受支持的前缀开头。";
            this.addQuote.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addQuote_ButtonClick);
            this.addQuote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addQuote_Click);
            this.addQuote.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addQuote_ItemsLoading);
            // 
            // show_refTaskPane
            // 
            this.show_refTaskPane.Image = global::PaperWriting.Properties.Resources.icons8_get_quote_48;
            this.show_refTaskPane.Label = "引用窗格";
            this.show_refTaskPane.Name = "show_refTaskPane";
            this.show_refTaskPane.ScreenTip = "打开引用窗格";
            this.show_refTaskPane.ShowImage = true;
            this.show_refTaskPane.SuperTip = "该任务窗格可更进一步方便您大量引用的操作。";
            // 
            // Main
            // 
            this.Name = "Main";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Main_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Insert.ResumeLayout(false);
            this.Insert.PerformLayout();
            this.quotes.ResumeLayout(false);
            this.quotes.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Insert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton inmaths;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup quotes;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox widthlimit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery addQuote;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton insert_figs;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton insert_label;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton from_file;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton from_clipboard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton tablelabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton figlabel;
        private Microsoft.Office.Tools.Ribbon.RibbonButton show_refTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery headers;
        private Microsoft.Office.Tools.Ribbon.RibbonButton h1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton h2;
        private Microsoft.Office.Tools.Ribbon.RibbonButton h3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton makequotable;
    }

    partial class ThisRibbonCollection
    {
        internal Main Main
        {
            get { return this.GetRibbon<Main>(); }
        }
    }
}
