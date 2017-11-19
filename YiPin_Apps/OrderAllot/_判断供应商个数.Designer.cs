namespace OrderAllot
{
    partial class _判断供应商个数
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
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.txtExport = new System.Windows.Forms.TextBox();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUpKsStore = new System.Windows.Forms.Button();
            this.btnUpShStore = new System.Windows.Forms.Button();
            this.txtUpKsStore = new System.Windows.Forms.TextBox();
            this.txtUpShStore = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpTime = new System.Windows.Forms.DateTimePicker();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dtpTime);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btnAnalyze);
            this.groupBox1.Controls.Add(this.txtExport);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnUpKsStore);
            this.groupBox1.Controls.Add(this.btnUpShStore);
            this.groupBox1.Controls.Add(this.txtUpKsStore);
            this.groupBox1.Controls.Add(this.txtUpShStore);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(368, 168);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(45, 109);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(35, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "结果:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 80);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 12);
            this.label5.TabIndex = 6;
            this.label5.Text = "昆山所有库存:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 53);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(83, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "上海所有库存:";
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Location = new System.Drawing.Point(288, 104);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 5;
            this.btnAnalyze.Text = "处理";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // txtExport
            // 
            this.txtExport.Location = new System.Drawing.Point(98, 104);
            this.txtExport.Name = "txtExport";
            this.txtExport.Size = new System.Drawing.Size(183, 21);
            this.txtExport.TabIndex = 4;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(93, 142);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 142);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "操作消息:";
            // 
            // btnUpKsStore
            // 
            this.btnUpKsStore.Location = new System.Drawing.Point(286, 75);
            this.btnUpKsStore.Name = "btnUpKsStore";
            this.btnUpKsStore.Size = new System.Drawing.Size(75, 23);
            this.btnUpKsStore.TabIndex = 1;
            this.btnUpKsStore.Text = "浏览";
            this.btnUpKsStore.UseVisualStyleBackColor = true;
            this.btnUpKsStore.Click += new System.EventHandler(this.btnUpKsStore_Click);
            // 
            // btnUpShStore
            // 
            this.btnUpShStore.Location = new System.Drawing.Point(287, 48);
            this.btnUpShStore.Name = "btnUpShStore";
            this.btnUpShStore.Size = new System.Drawing.Size(75, 23);
            this.btnUpShStore.TabIndex = 1;
            this.btnUpShStore.Text = "浏览";
            this.btnUpShStore.UseVisualStyleBackColor = true;
            this.btnUpShStore.Click += new System.EventHandler(this.btnUpShStore_Click);
            // 
            // txtUpKsStore
            // 
            this.txtUpKsStore.Enabled = false;
            this.txtUpKsStore.Location = new System.Drawing.Point(97, 77);
            this.txtUpKsStore.Name = "txtUpKsStore";
            this.txtUpKsStore.Size = new System.Drawing.Size(183, 21);
            this.txtUpKsStore.TabIndex = 0;
            // 
            // txtUpShStore
            // 
            this.txtUpShStore.Enabled = false;
            this.txtUpShStore.Location = new System.Drawing.Point(98, 50);
            this.txtUpShStore.Name = "txtUpShStore";
            this.txtUpShStore.Size = new System.Drawing.Size(183, 21);
            this.txtUpShStore.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "统计日期:";
            // 
            // dtpTime
            // 
            this.dtpTime.Location = new System.Drawing.Point(98, 21);
            this.dtpTime.Name = "dtpTime";
            this.dtpTime.Size = new System.Drawing.Size(200, 21);
            this.dtpTime.TabIndex = 7;
            // 
            // _判断供应商个数
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(387, 192);
            this.Controls.Add(this.groupBox1);
            this.Name = "_判断供应商个数";
            this.Text = "判断供应商个数";
            this.Load += new System.EventHandler(this._判断供应商个数_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.TextBox txtExport;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnUpKsStore;
        private System.Windows.Forms.Button btnUpShStore;
        private System.Windows.Forms.TextBox txtUpKsStore;
        private System.Windows.Forms.TextBox txtUpShStore;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpTime;
    }
}