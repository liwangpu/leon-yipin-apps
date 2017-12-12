namespace Gadget
{
    partial class _分库盘点
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
            this.txt盘点结果 = new System.Windows.Forms.TextBox();
            this.cb只要盘点结果 = new System.Windows.Forms.CheckBox();
            this.nd遗留上海天数 = new System.Windows.Forms.NumericUpDown();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.txtExport = new System.Windows.Forms.TextBox();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btn盘点结果 = new System.Windows.Forms.Button();
            this.btn库存明细 = new System.Windows.Forms.Button();
            this.txt库存明细 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nd遗留上海天数)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txt盘点结果);
            this.groupBox1.Controls.Add(this.cb只要盘点结果);
            this.groupBox1.Controls.Add(this.nd遗留上海天数);
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btnAnalyze);
            this.groupBox1.Controls.Add(this.txtExport);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btn盘点结果);
            this.groupBox1.Controls.Add(this.btn库存明细);
            this.groupBox1.Controls.Add(this.txt库存明细);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(357, 166);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // txt盘点结果
            // 
            this.txt盘点结果.Enabled = false;
            this.txt盘点结果.Location = new System.Drawing.Point(75, 48);
            this.txt盘点结果.Name = "txt盘点结果";
            this.txt盘点结果.Size = new System.Drawing.Size(183, 21);
            this.txt盘点结果.TabIndex = 23;
            // 
            // cb只要盘点结果
            // 
            this.cb只要盘点结果.AutoSize = true;
            this.cb只要盘点结果.Location = new System.Drawing.Point(243, 114);
            this.cb只要盘点结果.Name = "cb只要盘点结果";
            this.cb只要盘点结果.Size = new System.Drawing.Size(96, 16);
            this.cb只要盘点结果.TabIndex = 22;
            this.cb只要盘点结果.Text = "只要盘点结果";
            this.cb只要盘点结果.UseVisualStyleBackColor = true;
            // 
            // nd遗留上海天数
            // 
            this.nd遗留上海天数.Location = new System.Drawing.Point(108, 112);
            this.nd遗留上海天数.Name = "nd遗留上海天数";
            this.nd遗留上海天数.Size = new System.Drawing.Size(75, 21);
            this.nd遗留上海天数.TabIndex = 21;
            this.nd遗留上海天数.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(286, 145);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 20;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            this.lkDecs.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkDecs_LinkClicked);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(189, 114);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "天";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 115);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "预留(上海)天数:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 52);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "盘点结果:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "库存明细:";
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Location = new System.Drawing.Point(264, 79);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 5;
            this.btnAnalyze.Text = "处理";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // txtExport
            // 
            this.txtExport.Location = new System.Drawing.Point(75, 81);
            this.txtExport.Name = "txtExport";
            this.txtExport.Size = new System.Drawing.Size(183, 21);
            this.txtExport.TabIndex = 4;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(73, 145);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 145);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "操作消息:";
            // 
            // btn盘点结果
            // 
            this.btn盘点结果.Location = new System.Drawing.Point(264, 47);
            this.btn盘点结果.Name = "btn盘点结果";
            this.btn盘点结果.Size = new System.Drawing.Size(75, 23);
            this.btn盘点结果.TabIndex = 1;
            this.btn盘点结果.Text = "浏览";
            this.btn盘点结果.UseVisualStyleBackColor = true;
            this.btn盘点结果.Click += new System.EventHandler(this.btn盘点结果_Click);
            // 
            // btn库存明细
            // 
            this.btn库存明细.Location = new System.Drawing.Point(264, 18);
            this.btn库存明细.Name = "btn库存明细";
            this.btn库存明细.Size = new System.Drawing.Size(75, 23);
            this.btn库存明细.TabIndex = 1;
            this.btn库存明细.Text = "浏览";
            this.btn库存明细.UseVisualStyleBackColor = true;
            this.btn库存明细.Click += new System.EventHandler(this.btn库存明细_Click);
            // 
            // txt库存明细
            // 
            this.txt库存明细.Enabled = false;
            this.txt库存明细.Location = new System.Drawing.Point(75, 20);
            this.txt库存明细.Name = "txt库存明细";
            this.txt库存明细.Size = new System.Drawing.Size(183, 21);
            this.txt库存明细.TabIndex = 0;
            // 
            // _分库盘点
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(374, 184);
            this.Controls.Add(this.groupBox1);
            this.MaximumSize = new System.Drawing.Size(390, 222);
            this.MinimumSize = new System.Drawing.Size(390, 222);
            this.Name = "_分库盘点";
            this.Text = "_分库盘点";
            this.Load += new System.EventHandler(this._分库盘点_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nd遗留上海天数)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn库存明细;
        private System.Windows.Forms.TextBox txt库存明细;
        private System.Windows.Forms.TextBox txtExport;
        private System.Windows.Forms.Button btn盘点结果;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox cb只要盘点结果;
        private System.Windows.Forms.NumericUpDown nd遗留上海天数;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt盘点结果;



    }
}