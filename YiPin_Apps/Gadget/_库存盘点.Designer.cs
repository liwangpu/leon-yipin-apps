namespace Gadget
{
    partial class _库存盘点
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.txtExport = new System.Windows.Forms.TextBox();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnChuRuKu = new System.Windows.Forms.Button();
            this.btnUpKucun = new System.Windows.Forms.Button();
            this.btnUpJiaoHuo = new System.Windows.Forms.Button();
            this.txtChuRuKu = new System.Windows.Forms.TextBox();
            this.txtUpKucun = new System.Windows.Forms.TextBox();
            this.txtUpJiaoHuo = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.btnAnalyze);
            this.groupBox1.Controls.Add(this.txtExport);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnChuRuKu);
            this.groupBox1.Controls.Add(this.btnUpKucun);
            this.groupBox1.Controls.Add(this.btnUpJiaoHuo);
            this.groupBox1.Controls.Add(this.txtChuRuKu);
            this.groupBox1.Controls.Add(this.txtUpKucun);
            this.groupBox1.Controls.Add(this.txtUpJiaoHuo);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(349, 146);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(287, 127);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 7;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            this.lkDecs.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkDecs_LinkClicked);
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Location = new System.Drawing.Point(268, 98);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 5;
            this.btnAnalyze.Text = "处理";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // txtExport
            // 
            this.txtExport.Location = new System.Drawing.Point(78, 98);
            this.txtExport.Name = "txtExport";
            this.txtExport.Size = new System.Drawing.Size(183, 21);
            this.txtExport.TabIndex = 4;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(74, 127);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 101);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 12);
            this.label5.TabIndex = 2;
            this.label5.Text = "处理文件:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "出入库差";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "库存信息";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 17);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "拣货单:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 127);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "操作消息:";
            // 
            // btnChuRuKu
            // 
            this.btnChuRuKu.Location = new System.Drawing.Point(265, 69);
            this.btnChuRuKu.Name = "btnChuRuKu";
            this.btnChuRuKu.Size = new System.Drawing.Size(75, 23);
            this.btnChuRuKu.TabIndex = 1;
            this.btnChuRuKu.Text = "浏览";
            this.btnChuRuKu.UseVisualStyleBackColor = true;
            this.btnChuRuKu.Click += new System.EventHandler(this.btnChuRuKu_Click);
            // 
            // btnUpKucun
            // 
            this.btnUpKucun.Location = new System.Drawing.Point(265, 40);
            this.btnUpKucun.Name = "btnUpKucun";
            this.btnUpKucun.Size = new System.Drawing.Size(75, 23);
            this.btnUpKucun.TabIndex = 1;
            this.btnUpKucun.Text = "浏览";
            this.btnUpKucun.UseVisualStyleBackColor = true;
            this.btnUpKucun.Click += new System.EventHandler(this.btnUpKucun_Click);
            // 
            // btnUpJiaoHuo
            // 
            this.btnUpJiaoHuo.Location = new System.Drawing.Point(265, 12);
            this.btnUpJiaoHuo.Name = "btnUpJiaoHuo";
            this.btnUpJiaoHuo.Size = new System.Drawing.Size(75, 23);
            this.btnUpJiaoHuo.TabIndex = 1;
            this.btnUpJiaoHuo.Text = "浏览";
            this.btnUpJiaoHuo.UseVisualStyleBackColor = true;
            this.btnUpJiaoHuo.Click += new System.EventHandler(this.btnUpJiaoHuo_Click);
            // 
            // txtChuRuKu
            // 
            this.txtChuRuKu.Enabled = false;
            this.txtChuRuKu.Location = new System.Drawing.Point(76, 70);
            this.txtChuRuKu.Name = "txtChuRuKu";
            this.txtChuRuKu.Size = new System.Drawing.Size(183, 21);
            this.txtChuRuKu.TabIndex = 0;
            // 
            // txtUpKucun
            // 
            this.txtUpKucun.Enabled = false;
            this.txtUpKucun.Location = new System.Drawing.Point(76, 41);
            this.txtUpKucun.Name = "txtUpKucun";
            this.txtUpKucun.Size = new System.Drawing.Size(183, 21);
            this.txtUpKucun.TabIndex = 0;
            // 
            // txtUpJiaoHuo
            // 
            this.txtUpJiaoHuo.Enabled = false;
            this.txtUpJiaoHuo.Location = new System.Drawing.Point(76, 14);
            this.txtUpJiaoHuo.Name = "txtUpJiaoHuo";
            this.txtUpJiaoHuo.Size = new System.Drawing.Size(183, 21);
            this.txtUpJiaoHuo.TabIndex = 0;
            // 
            // _库存盘点
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 164);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(384, 202);
            this.MinimumSize = new System.Drawing.Size(384, 202);
            this.Name = "_库存盘点";
            this.Text = "库存盘点";
            this.Load += new System.EventHandler(this._库存盘点_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.TextBox txtExport;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnUpKucun;
        private System.Windows.Forms.Button btnUpJiaoHuo;
        private System.Windows.Forms.TextBox txtUpKucun;
        private System.Windows.Forms.TextBox txtUpJiaoHuo;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnChuRuKu;
        private System.Windows.Forms.TextBox txtChuRuKu;
    }
}