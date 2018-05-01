namespace Gadget
{
    partial class _配货绩效
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dtp绩效时间 = new System.Windows.Forms.DateTimePicker();
            this.lbMsg = new System.Windows.Forms.Label();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.label3 = new System.Windows.Forms.Label();
            this.btn库位人员配置 = new System.Windows.Forms.Button();
            this.btn全月绩效 = new System.Windows.Forms.Button();
            this.btn当天绩效 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn上传拣货时间 = new System.Windows.Forms.Button();
            this.txt拣货时间 = new System.Windows.Forms.TextBox();
            this.btn上传拣货单 = new System.Windows.Forms.Button();
            this.txt拣货单 = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lsbCache = new System.Windows.Forms.ListBox();
            this.cmn下载缓存 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.导出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.txt拣货人员配置 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.cmn下载缓存.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dtp绩效时间);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.btn库位人员配置);
            this.groupBox1.Controls.Add(this.btn全月绩效);
            this.groupBox1.Controls.Add(this.btn当天绩效);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btn上传拣货时间);
            this.groupBox1.Controls.Add(this.txt拣货时间);
            this.groupBox1.Controls.Add(this.btn上传拣货单);
            this.groupBox1.Controls.Add(this.txt拣货单);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(497, 131);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // dtp绩效时间
            // 
            this.dtp绩效时间.Location = new System.Drawing.Point(69, 78);
            this.dtp绩效时间.Name = "dtp绩效时间";
            this.dtp绩效时间.Size = new System.Drawing.Size(166, 21);
            this.dtp绩效时间.TabIndex = 14;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(69, 111);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 13;
            this.lbMsg.Text = "待上传文件";
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(391, 111);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 11;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            this.lkDecs.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkDecs_LinkClicked);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 111);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "操作消息:";
            // 
            // btn库位人员配置
            // 
            this.btn库位人员配置.Location = new System.Drawing.Point(241, 76);
            this.btn库位人员配置.Name = "btn库位人员配置";
            this.btn库位人员配置.Size = new System.Drawing.Size(88, 23);
            this.btn库位人员配置.TabIndex = 3;
            this.btn库位人员配置.Text = "库位人员配置";
            this.btn库位人员配置.UseVisualStyleBackColor = true;
            this.btn库位人员配置.Click += new System.EventHandler(this.btn库位人员配置_Click);
            // 
            // btn全月绩效
            // 
            this.btn全月绩效.Location = new System.Drawing.Point(335, 76);
            this.btn全月绩效.Name = "btn全月绩效";
            this.btn全月绩效.Size = new System.Drawing.Size(75, 23);
            this.btn全月绩效.TabIndex = 3;
            this.btn全月绩效.Text = "当月绩效";
            this.btn全月绩效.UseVisualStyleBackColor = true;
            this.btn全月绩效.Click += new System.EventHandler(this.btn全月绩效_Click);
            // 
            // btn当天绩效
            // 
            this.btn当天绩效.Enabled = false;
            this.btn当天绩效.Location = new System.Drawing.Point(416, 76);
            this.btn当天绩效.Name = "btn当天绩效";
            this.btn当天绩效.Size = new System.Drawing.Size(75, 23);
            this.btn当天绩效.TabIndex = 3;
            this.btn当天绩效.Text = "当天绩效";
            this.btn当天绩效.UseVisualStyleBackColor = true;
            this.btn当天绩效.Click += new System.EventHandler(this.btn当天绩效_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 84);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "绩效日期:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "拣货时间:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "拣货单:";
            // 
            // btn上传拣货时间
            // 
            this.btn上传拣货时间.Location = new System.Drawing.Point(416, 47);
            this.btn上传拣货时间.Name = "btn上传拣货时间";
            this.btn上传拣货时间.Size = new System.Drawing.Size(75, 23);
            this.btn上传拣货时间.TabIndex = 1;
            this.btn上传拣货时间.Text = "浏览";
            this.btn上传拣货时间.UseVisualStyleBackColor = true;
            this.btn上传拣货时间.Click += new System.EventHandler(this.btn上传拣货时间_Click);
            // 
            // txt拣货时间
            // 
            this.txt拣货时间.Enabled = false;
            this.txt拣货时间.Location = new System.Drawing.Point(71, 49);
            this.txt拣货时间.Name = "txt拣货时间";
            this.txt拣货时间.Size = new System.Drawing.Size(339, 21);
            this.txt拣货时间.TabIndex = 0;
            // 
            // btn上传拣货单
            // 
            this.btn上传拣货单.Location = new System.Drawing.Point(416, 18);
            this.btn上传拣货单.Name = "btn上传拣货单";
            this.btn上传拣货单.Size = new System.Drawing.Size(75, 23);
            this.btn上传拣货单.TabIndex = 1;
            this.btn上传拣货单.Text = "浏览";
            this.btn上传拣货单.UseVisualStyleBackColor = true;
            this.btn上传拣货单.Click += new System.EventHandler(this.btn上传拣货单_Click);
            // 
            // txt拣货单
            // 
            this.txt拣货单.Enabled = false;
            this.txt拣货单.Location = new System.Drawing.Point(71, 20);
            this.txt拣货单.Name = "txt拣货单";
            this.txt拣货单.Size = new System.Drawing.Size(339, 21);
            this.txt拣货单.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lsbCache);
            this.groupBox2.Location = new System.Drawing.Point(12, 149);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(497, 204);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "缓存信息";
            // 
            // lsbCache
            // 
            this.lsbCache.ContextMenuStrip = this.cmn下载缓存;
            this.lsbCache.FormattingEnabled = true;
            this.lsbCache.ItemHeight = 12;
            this.lsbCache.Location = new System.Drawing.Point(6, 20);
            this.lsbCache.Name = "lsbCache";
            this.lsbCache.Size = new System.Drawing.Size(485, 172);
            this.lsbCache.TabIndex = 0;
            // 
            // cmn下载缓存
            // 
            this.cmn下载缓存.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.导出ToolStripMenuItem});
            this.cmn下载缓存.Name = "cmn下载缓存";
            this.cmn下载缓存.Size = new System.Drawing.Size(101, 26);
            this.cmn下载缓存.Text = "导出";
            // 
            // 导出ToolStripMenuItem
            // 
            this.导出ToolStripMenuItem.Name = "导出ToolStripMenuItem";
            this.导出ToolStripMenuItem.Size = new System.Drawing.Size(100, 22);
            this.导出ToolStripMenuItem.Text = "导出";
            this.导出ToolStripMenuItem.Click += new System.EventHandler(this.导出ToolStripMenuItem_Click);
            // 
            // txt拣货人员配置
            // 
            this.txt拣货人员配置.Location = new System.Drawing.Point(455, 3);
            this.txt拣货人员配置.Name = "txt拣货人员配置";
            this.txt拣货人员配置.Size = new System.Drawing.Size(14, 21);
            this.txt拣货人员配置.TabIndex = 9;
            this.txt拣货人员配置.Visible = false;
            // 
            // _配货绩效
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 360);
            this.Controls.Add(this.txt拣货人员配置);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(537, 399);
            this.MinimumSize = new System.Drawing.Size(537, 399);
            this.Name = "_配货绩效";
            this.Text = "配货绩效";
            this.Load += new System.EventHandler(this._配货绩效_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.cmn下载缓存.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn上传拣货单;
        private System.Windows.Forms.TextBox txt拣货单;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListBox lsbCache;
        private System.Windows.Forms.Button btn当天绩效;
        private System.Windows.Forms.Button btn全月绩效;
        private System.Windows.Forms.Button btn库位人员配置;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn上传拣货时间;
        private System.Windows.Forms.TextBox txt拣货时间;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.TextBox txt拣货人员配置;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtp绩效时间;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ContextMenuStrip cmn下载缓存;
        private System.Windows.Forms.ToolStripMenuItem 导出ToolStripMenuItem;
    }
}