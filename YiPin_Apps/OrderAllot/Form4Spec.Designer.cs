namespace OrderAllot
{
    partial class Form4Spec
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
            this.NtxtAmount = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.txtExport = new System.Windows.Forms.TextBox();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUpTmp = new System.Windows.Forms.Button();
            this.btnUpSHKc = new System.Windows.Forms.Button();
            this.btnKsKc = new System.Windows.Forms.Button();
            this.btnUpKsYj = new System.Windows.Forms.Button();
            this.txtUpTmp = new System.Windows.Forms.TextBox();
            this.btnUpDfkunsYj = new System.Windows.Forms.Button();
            this.txtUpSHKc = new System.Windows.Forms.TextBox();
            this.txtUpKsKc = new System.Windows.Forms.TextBox();
            this.txtUpKsYj = new System.Windows.Forms.TextBox();
            this.txtUpDfkunsYj = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NtxtAmount)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.NtxtAmount);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.btnAnalyze);
            this.groupBox1.Controls.Add(this.txtExport);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnUpTmp);
            this.groupBox1.Controls.Add(this.btnUpSHKc);
            this.groupBox1.Controls.Add(this.btnKsKc);
            this.groupBox1.Controls.Add(this.btnUpKsYj);
            this.groupBox1.Controls.Add(this.txtUpTmp);
            this.groupBox1.Controls.Add(this.btnUpDfkunsYj);
            this.groupBox1.Controls.Add(this.txtUpSHKc);
            this.groupBox1.Controls.Add(this.txtUpKsKc);
            this.groupBox1.Controls.Add(this.txtUpKsYj);
            this.groupBox1.Controls.Add(this.txtUpDfkunsYj);
            this.groupBox1.Location = new System.Drawing.Point(11, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(385, 246);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // NtxtAmount
            // 
            this.NtxtAmount.Location = new System.Drawing.Point(73, 18);
            this.NtxtAmount.Maximum = new decimal(new int[] {
            9999999,
            0,
            0,
            0});
            this.NtxtAmount.Name = "NtxtAmount";
            this.NtxtAmount.Size = new System.Drawing.Size(99, 21);
            this.NtxtAmount.TabIndex = 10;
            this.NtxtAmount.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(179, 20);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(17, 12);
            this.label6.TabIndex = 8;
            this.label6.Text = "元";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 20);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 12);
            this.label7.TabIndex = 9;
            this.label7.Text = "订单金额:";
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Location = new System.Drawing.Point(303, 178);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 5;
            this.btnAnalyze.Text = "处理";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // txtExport
            // 
            this.txtExport.Location = new System.Drawing.Point(113, 178);
            this.txtExport.Name = "txtExport";
            this.txtExport.Size = new System.Drawing.Size(183, 21);
            this.txtExport.TabIndex = 4;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(70, 221);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(5, 181);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 12);
            this.label5.TabIndex = 2;
            this.label5.Text = "处理文件:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(5, 154);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(59, 12);
            this.label9.TabIndex = 2;
            this.label9.Text = "临时备货:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(5, 130);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(83, 12);
            this.label8.TabIndex = 2;
            this.label8.Text = "上海所有库存:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 103);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "昆山所有库存:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 76);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "昆山采购建议:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "默认昆山预警订单:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 221);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "操作消息:";
            // 
            // btnUpTmp
            // 
            this.btnUpTmp.Location = new System.Drawing.Point(303, 149);
            this.btnUpTmp.Name = "btnUpTmp";
            this.btnUpTmp.Size = new System.Drawing.Size(75, 23);
            this.btnUpTmp.TabIndex = 1;
            this.btnUpTmp.Text = "浏览";
            this.btnUpTmp.UseVisualStyleBackColor = true;
            this.btnUpTmp.Click += new System.EventHandler(this.btnUpTmp_Click);
            // 
            // btnUpSHKc
            // 
            this.btnUpSHKc.Location = new System.Drawing.Point(303, 125);
            this.btnUpSHKc.Name = "btnUpSHKc";
            this.btnUpSHKc.Size = new System.Drawing.Size(75, 23);
            this.btnUpSHKc.TabIndex = 1;
            this.btnUpSHKc.Text = "浏览";
            this.btnUpSHKc.UseVisualStyleBackColor = true;
            this.btnUpSHKc.Click += new System.EventHandler(this.btnUpSHKc_Click);
            // 
            // btnKsKc
            // 
            this.btnKsKc.Location = new System.Drawing.Point(304, 98);
            this.btnKsKc.Name = "btnKsKc";
            this.btnKsKc.Size = new System.Drawing.Size(75, 23);
            this.btnKsKc.TabIndex = 1;
            this.btnKsKc.Text = "浏览";
            this.btnKsKc.UseVisualStyleBackColor = true;
            this.btnKsKc.Click += new System.EventHandler(this.btnKsKc_Click);
            // 
            // btnUpKsYj
            // 
            this.btnUpKsYj.Location = new System.Drawing.Point(304, 71);
            this.btnUpKsYj.Name = "btnUpKsYj";
            this.btnUpKsYj.Size = new System.Drawing.Size(75, 23);
            this.btnUpKsYj.TabIndex = 1;
            this.btnUpKsYj.Text = "浏览";
            this.btnUpKsYj.UseVisualStyleBackColor = true;
            this.btnUpKsYj.Click += new System.EventHandler(this.btnUpKsYj_Click);
            // 
            // txtUpTmp
            // 
            this.txtUpTmp.Enabled = false;
            this.txtUpTmp.Location = new System.Drawing.Point(114, 150);
            this.txtUpTmp.Name = "txtUpTmp";
            this.txtUpTmp.Size = new System.Drawing.Size(183, 21);
            this.txtUpTmp.TabIndex = 0;
            // 
            // btnUpDfkunsYj
            // 
            this.btnUpDfkunsYj.Location = new System.Drawing.Point(304, 43);
            this.btnUpDfkunsYj.Name = "btnUpDfkunsYj";
            this.btnUpDfkunsYj.Size = new System.Drawing.Size(75, 23);
            this.btnUpDfkunsYj.TabIndex = 1;
            this.btnUpDfkunsYj.Text = "浏览";
            this.btnUpDfkunsYj.UseVisualStyleBackColor = true;
            this.btnUpDfkunsYj.Click += new System.EventHandler(this.btnUpDfkunsYj_Click);
            // 
            // txtUpSHKc
            // 
            this.txtUpSHKc.Enabled = false;
            this.txtUpSHKc.Location = new System.Drawing.Point(114, 126);
            this.txtUpSHKc.Name = "txtUpSHKc";
            this.txtUpSHKc.Size = new System.Drawing.Size(183, 21);
            this.txtUpSHKc.TabIndex = 0;
            // 
            // txtUpKsKc
            // 
            this.txtUpKsKc.Enabled = false;
            this.txtUpKsKc.Location = new System.Drawing.Point(115, 99);
            this.txtUpKsKc.Name = "txtUpKsKc";
            this.txtUpKsKc.Size = new System.Drawing.Size(183, 21);
            this.txtUpKsKc.TabIndex = 0;
            // 
            // txtUpKsYj
            // 
            this.txtUpKsYj.Enabled = false;
            this.txtUpKsYj.Location = new System.Drawing.Point(115, 72);
            this.txtUpKsYj.Name = "txtUpKsYj";
            this.txtUpKsYj.Size = new System.Drawing.Size(183, 21);
            this.txtUpKsYj.TabIndex = 0;
            // 
            // txtUpDfkunsYj
            // 
            this.txtUpDfkunsYj.Enabled = false;
            this.txtUpDfkunsYj.Location = new System.Drawing.Point(115, 45);
            this.txtUpDfkunsYj.Name = "txtUpDfkunsYj";
            this.txtUpDfkunsYj.Size = new System.Drawing.Size(183, 21);
            this.txtUpDfkunsYj.TabIndex = 0;
            // 
            // Form4Spec
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(408, 263);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form4Spec";
            this.Text = "订单分配(排除重复项)-临时备货";
            this.Load += new System.EventHandler(this.Form4Spec_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NtxtAmount)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.TextBox txtExport;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnKsKc;
        private System.Windows.Forms.Button btnUpKsYj;
        private System.Windows.Forms.Button btnUpDfkunsYj;
        private System.Windows.Forms.TextBox txtUpKsKc;
        private System.Windows.Forms.TextBox txtUpKsYj;
        private System.Windows.Forms.TextBox txtUpDfkunsYj;
        private System.Windows.Forms.NumericUpDown NtxtAmount;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnUpSHKc;
        private System.Windows.Forms.TextBox txtUpSHKc;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnUpTmp;
        private System.Windows.Forms.TextBox txtUpTmp;

    }
}