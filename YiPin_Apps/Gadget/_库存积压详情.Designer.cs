namespace Gadget
{
    partial class _库存积压详情
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
            this.ndp积压天数 = new System.Windows.Forms.NumericUpDown();
            this.ndp可用数量 = new System.Windows.Forms.NumericUpDown();
            this.ndp库存金额 = new System.Windows.Forms.NumericUpDown();
            this.ndp周转天数 = new System.Windows.Forms.NumericUpDown();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.label7 = new System.Windows.Forms.Label();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btn处理 = new System.Windows.Forms.Button();
            this.btn上传入库明细 = new System.Windows.Forms.Button();
            this.btn上传库存周转率 = new System.Windows.Forms.Button();
            this.txt入库明细表 = new System.Windows.Forms.TextBox();
            this.txt库存周转率 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ndp积压天数)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ndp可用数量)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ndp库存金额)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ndp周转天数)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ndp积压天数);
            this.groupBox1.Controls.Add(this.ndp可用数量);
            this.groupBox1.Controls.Add(this.ndp库存金额);
            this.groupBox1.Controls.Add(this.ndp周转天数);
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btn处理);
            this.groupBox1.Controls.Add(this.btn上传入库明细);
            this.groupBox1.Controls.Add(this.btn上传库存周转率);
            this.groupBox1.Controls.Add(this.txt入库明细表);
            this.groupBox1.Controls.Add(this.txt库存周转率);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(365, 219);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // ndp积压天数
            // 
            this.ndp积压天数.Location = new System.Drawing.Point(86, 156);
            this.ndp积压天数.Maximum = new decimal(new int[] {
            9999999,
            0,
            0,
            0});
            this.ndp积压天数.Name = "ndp积压天数";
            this.ndp积压天数.Size = new System.Drawing.Size(192, 21);
            this.ndp积压天数.TabIndex = 24;
            // 
            // ndp可用数量
            // 
            this.ndp可用数量.Location = new System.Drawing.Point(86, 129);
            this.ndp可用数量.Maximum = new decimal(new int[] {
            9999999,
            0,
            0,
            0});
            this.ndp可用数量.Name = "ndp可用数量";
            this.ndp可用数量.Size = new System.Drawing.Size(192, 21);
            this.ndp可用数量.TabIndex = 24;
            // 
            // ndp库存金额
            // 
            this.ndp库存金额.Location = new System.Drawing.Point(86, 102);
            this.ndp库存金额.Maximum = new decimal(new int[] {
            9999999,
            0,
            0,
            0});
            this.ndp库存金额.Name = "ndp库存金额";
            this.ndp库存金额.Size = new System.Drawing.Size(192, 21);
            this.ndp库存金额.TabIndex = 24;
            // 
            // ndp周转天数
            // 
            this.ndp周转天数.Location = new System.Drawing.Point(86, 75);
            this.ndp周转天数.Maximum = new decimal(new int[] {
            9999999,
            0,
            0,
            0});
            this.ndp周转天数.Name = "ndp周转天数";
            this.ndp周转天数.Size = new System.Drawing.Size(192, 21);
            this.ndp周转天数.TabIndex = 24;
            this.ndp周转天数.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(304, 193);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 23;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            this.lkDecs.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkDecs_LinkClicked);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(19, 160);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 12);
            this.label7.TabIndex = 9;
            this.label7.Text = "积压天数";
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(93, 193);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 22;
            this.lbMsg.Text = "待上传文件";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(19, 133);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 9;
            this.label6.Text = "可用数量";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(19, 106);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 9;
            this.label5.Text = "库存金额";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(27, 193);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 21;
            this.label3.Text = "操作消息:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 9;
            this.label2.Text = "周转天数:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 12);
            this.label1.TabIndex = 9;
            this.label1.Text = "入库明细表:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 12);
            this.label4.TabIndex = 9;
            this.label4.Text = "库存周转率:";
            // 
            // btn处理
            // 
            this.btn处理.Location = new System.Drawing.Point(284, 74);
            this.btn处理.Name = "btn处理";
            this.btn处理.Size = new System.Drawing.Size(75, 103);
            this.btn处理.TabIndex = 8;
            this.btn处理.Text = "处理";
            this.btn处理.UseVisualStyleBackColor = true;
            this.btn处理.Click += new System.EventHandler(this.btn处理_Click);
            // 
            // btn上传入库明细
            // 
            this.btn上传入库明细.Location = new System.Drawing.Point(284, 45);
            this.btn上传入库明细.Name = "btn上传入库明细";
            this.btn上传入库明细.Size = new System.Drawing.Size(75, 23);
            this.btn上传入库明细.TabIndex = 8;
            this.btn上传入库明细.Text = "浏览";
            this.btn上传入库明细.UseVisualStyleBackColor = true;
            this.btn上传入库明细.Click += new System.EventHandler(this.btn上传入库明细_Click);
            // 
            // btn上传库存周转率
            // 
            this.btn上传库存周转率.Location = new System.Drawing.Point(284, 18);
            this.btn上传库存周转率.Name = "btn上传库存周转率";
            this.btn上传库存周转率.Size = new System.Drawing.Size(75, 23);
            this.btn上传库存周转率.TabIndex = 8;
            this.btn上传库存周转率.Text = "浏览";
            this.btn上传库存周转率.UseVisualStyleBackColor = true;
            this.btn上传库存周转率.Click += new System.EventHandler(this.btn上传库存周转率_Click);
            // 
            // txt入库明细表
            // 
            this.txt入库明细表.Enabled = false;
            this.txt入库明细表.Location = new System.Drawing.Point(86, 47);
            this.txt入库明细表.Name = "txt入库明细表";
            this.txt入库明细表.Size = new System.Drawing.Size(192, 21);
            this.txt入库明细表.TabIndex = 7;
            // 
            // txt库存周转率
            // 
            this.txt库存周转率.Enabled = false;
            this.txt库存周转率.Location = new System.Drawing.Point(86, 18);
            this.txt库存周转率.Name = "txt库存周转率";
            this.txt库存周转率.Size = new System.Drawing.Size(192, 21);
            this.txt库存周转率.TabIndex = 7;
            // 
            // _库存积压详情
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(383, 235);
            this.Controls.Add(this.groupBox1);
            this.MaximumSize = new System.Drawing.Size(399, 273);
            this.MinimumSize = new System.Drawing.Size(399, 273);
            this.Name = "_库存积压详情";
            this.Text = "_库存积压详情";
            this.Load += new System.EventHandler(this._库存积压详情_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ndp积压天数)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ndp可用数量)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ndp库存金额)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ndp周转天数)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btn处理;
        private System.Windows.Forms.Button btn上传入库明细;
        private System.Windows.Forms.Button btn上传库存周转率;
        private System.Windows.Forms.TextBox txt入库明细表;
        private System.Windows.Forms.TextBox txt库存周转率;
        private System.Windows.Forms.NumericUpDown ndp周转天数;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown ndp可用数量;
        private System.Windows.Forms.NumericUpDown ndp库存金额;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown ndp积压天数;
        private System.Windows.Forms.Label label7;
    }
}