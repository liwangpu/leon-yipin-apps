namespace OrderAllot
{
    partial class Form1
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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cb30day = new System.Windows.Forms.CheckBox();
            this.cb15day = new System.Windows.Forms.CheckBox();
            this.cb5day = new System.Windows.Forms.CheckBox();
            this.label10 = new System.Windows.Forms.Label();
            this.btnUpHot = new System.Windows.Forms.Button();
            this.txtUpHot = new System.Windows.Forms.TextBox();
            this.NtxtAmount = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.txtExport = new System.Windows.Forms.TextBox();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUpKunsStore = new System.Windows.Forms.Button();
            this.btnUpShangsYj = new System.Windows.Forms.Button();
            this.txtUpKunsStore = new System.Windows.Forms.TextBox();
            this.txtUpShangsYj = new System.Windows.Forms.TextBox();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NtxtAmount)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.cb30day);
            this.groupBox1.Controls.Add(this.cb15day);
            this.groupBox1.Controls.Add(this.cb5day);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.btnUpHot);
            this.groupBox1.Controls.Add(this.txtUpHot);
            this.groupBox1.Controls.Add(this.NtxtAmount);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnAnalyze);
            this.groupBox1.Controls.Add(this.txtExport);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnUpKunsStore);
            this.groupBox1.Controls.Add(this.btnUpShangsYj);
            this.groupBox1.Controls.Add(this.txtUpKunsStore);
            this.groupBox1.Controls.Add(this.txtUpShangsYj);
            this.groupBox1.Location = new System.Drawing.Point(12, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(368, 225);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // cb30day
            // 
            this.cb30day.AutoSize = true;
            this.cb30day.Checked = true;
            this.cb30day.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cb30day.Location = new System.Drawing.Point(212, 141);
            this.cb30day.Name = "cb30day";
            this.cb30day.Size = new System.Drawing.Size(48, 16);
            this.cb30day.TabIndex = 17;
            this.cb30day.Text = "30天";
            this.cb30day.UseVisualStyleBackColor = true;
            // 
            // cb15day
            // 
            this.cb15day.AutoSize = true;
            this.cb15day.Checked = true;
            this.cb15day.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cb15day.Location = new System.Drawing.Point(158, 141);
            this.cb15day.Name = "cb15day";
            this.cb15day.Size = new System.Drawing.Size(48, 16);
            this.cb15day.TabIndex = 18;
            this.cb15day.Text = "15天";
            this.cb15day.UseVisualStyleBackColor = true;
            // 
            // cb5day
            // 
            this.cb5day.AutoSize = true;
            this.cb5day.Checked = true;
            this.cb5day.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cb5day.Location = new System.Drawing.Point(110, 141);
            this.cb5day.Name = "cb5day";
            this.cb5day.Size = new System.Drawing.Size(42, 16);
            this.cb5day.TabIndex = 19;
            this.cb5day.Text = "5天";
            this.cb5day.UseVisualStyleBackColor = true;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(31, 115);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(59, 12);
            this.label10.TabIndex = 16;
            this.label10.Text = "热销订单:";
            // 
            // btnUpHot
            // 
            this.btnUpHot.Location = new System.Drawing.Point(286, 111);
            this.btnUpHot.Name = "btnUpHot";
            this.btnUpHot.Size = new System.Drawing.Size(75, 23);
            this.btnUpHot.TabIndex = 13;
            this.btnUpHot.Text = "浏览";
            this.btnUpHot.UseVisualStyleBackColor = true;
            this.btnUpHot.Click += new System.EventHandler(this.btnUpHot_Click);
            // 
            // txtUpHot
            // 
            this.txtUpHot.Enabled = false;
            this.txtUpHot.Location = new System.Drawing.Point(97, 112);
            this.txtUpHot.Name = "txtUpHot";
            this.txtUpHot.Size = new System.Drawing.Size(183, 21);
            this.txtUpHot.TabIndex = 12;
            // 
            // NtxtAmount
            // 
            this.NtxtAmount.Location = new System.Drawing.Point(97, 19);
            this.NtxtAmount.Maximum = new decimal(new int[] {
            9999999,
            0,
            0,
            0});
            this.NtxtAmount.Name = "NtxtAmount";
            this.NtxtAmount.Size = new System.Drawing.Size(99, 21);
            this.NtxtAmount.TabIndex = 7;
            this.NtxtAmount.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(179, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "元";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(42, 163);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(35, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "结果:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 88);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 12);
            this.label5.TabIndex = 6;
            this.label5.Text = "昆山所有库存:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 61);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "上海库存预警:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(32, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "订单金额:";
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Location = new System.Drawing.Point(286, 163);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 5;
            this.btnAnalyze.Text = "处理";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // txtExport
            // 
            this.txtExport.Location = new System.Drawing.Point(96, 163);
            this.txtExport.Name = "txtExport";
            this.txtExport.Size = new System.Drawing.Size(183, 21);
            this.txtExport.TabIndex = 4;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(94, 196);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 196);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "操作消息:";
            // 
            // btnUpKunsStore
            // 
            this.btnUpKunsStore.Location = new System.Drawing.Point(285, 83);
            this.btnUpKunsStore.Name = "btnUpKunsStore";
            this.btnUpKunsStore.Size = new System.Drawing.Size(75, 23);
            this.btnUpKunsStore.TabIndex = 1;
            this.btnUpKunsStore.Text = "浏览";
            this.btnUpKunsStore.UseVisualStyleBackColor = true;
            this.btnUpKunsStore.Click += new System.EventHandler(this.btnUpKunsStore_Click);
            // 
            // btnUpShangsYj
            // 
            this.btnUpShangsYj.Location = new System.Drawing.Point(286, 56);
            this.btnUpShangsYj.Name = "btnUpShangsYj";
            this.btnUpShangsYj.Size = new System.Drawing.Size(75, 23);
            this.btnUpShangsYj.TabIndex = 1;
            this.btnUpShangsYj.Text = "浏览";
            this.btnUpShangsYj.UseVisualStyleBackColor = true;
            this.btnUpShangsYj.Click += new System.EventHandler(this.btnUpShangsYj_Click);
            // 
            // txtUpKunsStore
            // 
            this.txtUpKunsStore.Enabled = false;
            this.txtUpKunsStore.Location = new System.Drawing.Point(96, 85);
            this.txtUpKunsStore.Name = "txtUpKunsStore";
            this.txtUpKunsStore.Size = new System.Drawing.Size(183, 21);
            this.txtUpKunsStore.TabIndex = 0;
            // 
            // txtUpShangsYj
            // 
            this.txtUpShangsYj.Enabled = false;
            this.txtUpShangsYj.Location = new System.Drawing.Point(97, 58);
            this.txtUpShangsYj.Name = "txtUpShangsYj";
            this.txtUpShangsYj.Size = new System.Drawing.Size(183, 21);
            this.txtUpShangsYj.TabIndex = 0;
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(307, 196);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 20;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            this.lkDecs.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkDecs_LinkClicked);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(392, 241);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(408, 279);
            this.MinimumSize = new System.Drawing.Size(408, 279);
            this.Name = "Form1";
            this.Text = "订单分配";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NtxtAmount)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnUpShangsYj;
        private System.Windows.Forms.TextBox txtUpShangsYj;
        private System.Windows.Forms.TextBox txtExport;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown NtxtAmount;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnUpKunsStore;
        private System.Windows.Forms.TextBox txtUpKunsStore;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btnUpHot;
        private System.Windows.Forms.TextBox txtUpHot;
        private System.Windows.Forms.CheckBox cb30day;
        private System.Windows.Forms.CheckBox cb15day;
        private System.Windows.Forms.CheckBox cb5day;
        private System.Windows.Forms.LinkLabel lkDecs;
    }
}

