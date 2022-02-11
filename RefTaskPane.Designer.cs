namespace PaperWriting
{
    partial class RefTaskPane
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.ListViewGroup listViewGroup1 = new System.Windows.Forms.ListViewGroup("公式", System.Windows.Forms.HorizontalAlignment.Left);
            System.Windows.Forms.ListViewGroup listViewGroup2 = new System.Windows.Forms.ListViewGroup("图片", System.Windows.Forms.HorizontalAlignment.Left);
            System.Windows.Forms.ListViewGroup listViewGroup3 = new System.Windows.Forms.ListViewGroup("表格", System.Windows.Forms.HorizontalAlignment.Left);
            this.contextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.引用选中项ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除选中项ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.刷新ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.imageList = new System.Windows.Forms.ImageList(this.components);
            this.delete = new System.Windows.Forms.Button();
            this.insert = new System.Windows.Forms.Button();
            this.refresh = new System.Windows.Forms.Button();
            this.hyperref = new System.Windows.Forms.CheckBox();
            this.QuotableContents = new System.Windows.Forms.ListView();
            this.contextMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenu
            // 
            this.contextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.引用选中项ToolStripMenuItem,
            this.删除选中项ToolStripMenuItem,
            this.toolStripSeparator1,
            this.刷新ToolStripMenuItem});
            this.contextMenu.Name = "contextMenu";
            this.contextMenu.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.contextMenu.ShowImageMargin = false;
            this.contextMenu.Size = new System.Drawing.Size(173, 76);
            // 
            // 引用选中项ToolStripMenuItem
            // 
            this.引用选中项ToolStripMenuItem.Name = "引用选中项ToolStripMenuItem";
            this.引用选中项ToolStripMenuItem.ShortcutKeyDisplayString = "空格/回车";
            this.引用选中项ToolStripMenuItem.Size = new System.Drawing.Size(172, 22);
            this.引用选中项ToolStripMenuItem.Text = "引用选中项";
            this.引用选中项ToolStripMenuItem.Click += new System.EventHandler(this.引用选中项ToolStripMenuItem_Click);
            // 
            // 删除选中项ToolStripMenuItem
            // 
            this.删除选中项ToolStripMenuItem.Name = "删除选中项ToolStripMenuItem";
            this.删除选中项ToolStripMenuItem.ShortcutKeyDisplayString = "Del";
            this.删除选中项ToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.Delete;
            this.删除选中项ToolStripMenuItem.Size = new System.Drawing.Size(172, 22);
            this.删除选中项ToolStripMenuItem.Text = "删除选中项";
            this.删除选中项ToolStripMenuItem.Click += new System.EventHandler(this.删除选中项ToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(169, 6);
            // 
            // 刷新ToolStripMenuItem
            // 
            this.刷新ToolStripMenuItem.Name = "刷新ToolStripMenuItem";
            this.刷新ToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F5;
            this.刷新ToolStripMenuItem.Size = new System.Drawing.Size(172, 22);
            this.刷新ToolStripMenuItem.Text = "刷新";
            this.刷新ToolStripMenuItem.Click += new System.EventHandler(this.刷新ToolStripMenuItem_Click);
            // 
            // imageList
            // 
            this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth16Bit;
            this.imageList.ImageSize = new System.Drawing.Size(256, 64);
            this.imageList.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // delete
            // 
            this.delete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.delete.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(253)))), ((int)(((byte)(253)))));
            this.delete.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(171)))), ((int)(((byte)(171)))), ((int)(((byte)(171)))));
            this.delete.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(189)))), ((int)(((byte)(227)))));
            this.delete.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(213)))), ((int)(((byte)(225)))), ((int)(((byte)(242)))));
            this.delete.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.delete.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.delete.Image = global::PaperWriting.Properties.Resources.trash;
            this.delete.Location = new System.Drawing.Point(265, 467);
            this.delete.Name = "delete";
            this.delete.Size = new System.Drawing.Size(62, 28);
            this.delete.TabIndex = 10;
            this.delete.Text = "删除";
            this.delete.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.delete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.delete.UseVisualStyleBackColor = true;
            this.delete.Click += new System.EventHandler(this.delete_Click);
            this.delete.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btn_MouseDown);
            this.delete.MouseEnter += new System.EventHandler(this.btn_MouseEnter);
            this.delete.MouseLeave += new System.EventHandler(this.btn_MouseLeave);
            this.delete.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btn_MouseUp);
            // 
            // insert
            // 
            this.insert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.insert.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(253)))), ((int)(((byte)(253)))));
            this.insert.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(171)))), ((int)(((byte)(171)))), ((int)(((byte)(171)))));
            this.insert.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(189)))), ((int)(((byte)(227)))));
            this.insert.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(213)))), ((int)(((byte)(225)))), ((int)(((byte)(242)))));
            this.insert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.insert.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.insert.Image = global::PaperWriting.Properties.Resources.link;
            this.insert.Location = new System.Drawing.Point(197, 467);
            this.insert.Name = "insert";
            this.insert.Size = new System.Drawing.Size(62, 28);
            this.insert.TabIndex = 9;
            this.insert.Text = "引用";
            this.insert.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.insert.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.insert.UseVisualStyleBackColor = true;
            this.insert.Click += new System.EventHandler(this.insert_Click);
            this.insert.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btn_MouseDown);
            this.insert.MouseEnter += new System.EventHandler(this.btn_MouseEnter);
            this.insert.MouseLeave += new System.EventHandler(this.btn_MouseLeave);
            this.insert.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btn_MouseUp);
            // 
            // refresh
            // 
            this.refresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.refresh.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(253)))), ((int)(((byte)(253)))));
            this.refresh.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(171)))), ((int)(((byte)(171)))), ((int)(((byte)(171)))));
            this.refresh.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(189)))), ((int)(((byte)(227)))));
            this.refresh.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(213)))), ((int)(((byte)(225)))), ((int)(((byte)(242)))));
            this.refresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.refresh.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.refresh.Image = global::PaperWriting.Properties.Resources.refresh;
            this.refresh.Location = new System.Drawing.Point(129, 467);
            this.refresh.Name = "refresh";
            this.refresh.Size = new System.Drawing.Size(62, 28);
            this.refresh.TabIndex = 8;
            this.refresh.Text = "刷新";
            this.refresh.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.refresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.refresh.UseVisualStyleBackColor = true;
            this.refresh.Click += new System.EventHandler(this.refresh_Click);
            this.refresh.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btn_MouseDown);
            this.refresh.MouseEnter += new System.EventHandler(this.btn_MouseEnter);
            this.refresh.MouseLeave += new System.EventHandler(this.btn_MouseLeave);
            this.refresh.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btn_MouseUp);
            // 
            // hyperref
            // 
            this.hyperref.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.hyperref.AutoSize = true;
            this.hyperref.Checked = true;
            this.hyperref.CheckState = System.Windows.Forms.CheckState.Checked;
            this.hyperref.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(171)))), ((int)(((byte)(171)))), ((int)(((byte)(171)))));
            this.hyperref.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(189)))), ((int)(((byte)(227)))));
            this.hyperref.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(213)))), ((int)(((byte)(225)))), ((int)(((byte)(242)))));
            this.hyperref.Location = new System.Drawing.Point(10, 472);
            this.hyperref.Margin = new System.Windows.Forms.Padding(4);
            this.hyperref.Name = "hyperref";
            this.hyperref.Size = new System.Drawing.Size(107, 21);
            this.hyperref.TabIndex = 7;
            this.hyperref.Text = "链接到内容";
            this.hyperref.UseVisualStyleBackColor = true;
            // 
            // QuotableContents
            // 
            this.QuotableContents.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.QuotableContents.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.QuotableContents.ContextMenuStrip = this.contextMenu;
            listViewGroup1.Header = "公式";
            listViewGroup1.Name = "公式";
            listViewGroup2.Header = "图片";
            listViewGroup2.Name = "图片";
            listViewGroup3.Header = "表格";
            listViewGroup3.Name = "表格";
            this.QuotableContents.Groups.AddRange(new System.Windows.Forms.ListViewGroup[] {
            listViewGroup1,
            listViewGroup2,
            listViewGroup3});
            this.QuotableContents.HideSelection = false;
            this.QuotableContents.LargeImageList = this.imageList;
            this.QuotableContents.Location = new System.Drawing.Point(10, 10);
            this.QuotableContents.Name = "QuotableContents";
            this.QuotableContents.Size = new System.Drawing.Size(320, 447);
            this.QuotableContents.TabIndex = 6;
            this.QuotableContents.UseCompatibleStateImageBehavior = false;
            this.QuotableContents.DoubleClick += new System.EventHandler(this.QuotableContents_DoubleClick);
            // 
            // RefTaskPane
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.Controls.Add(this.delete);
            this.Controls.Add(this.insert);
            this.Controls.Add(this.refresh);
            this.Controls.Add(this.hyperref);
            this.Controls.Add(this.QuotableContents);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MinimumSize = new System.Drawing.Size(340, 0);
            this.Name = "RefTaskPane";
            this.Padding = new System.Windows.Forms.Padding(10);
            this.Size = new System.Drawing.Size(340, 514);
            this.VisibleChanged += new System.EventHandler(this.RefTaskPane_VisibleChanged);
            this.contextMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ImageList imageList;
        private System.Windows.Forms.ContextMenuStrip contextMenu;
        private System.Windows.Forms.ToolStripMenuItem 引用选中项ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除选中项ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem 刷新ToolStripMenuItem;
        private System.Windows.Forms.Button delete;
        private System.Windows.Forms.Button insert;
        private System.Windows.Forms.Button refresh;
        private System.Windows.Forms.CheckBox hyperref;
        private System.Windows.Forms.ListView QuotableContents;
    }
}
