﻿namespace OrderAllot
{
    partial class Form5
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
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.txtExport = new System.Windows.Forms.TextBox();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUpShRp = new System.Windows.Forms.Button();
            this.txtUpShRp = new System.Windows.Forms.TextBox();
            this.txtUpKsRp = new System.Windows.Forms.TextBox();
            this.btnUpKsRp = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnUpKsRp);
            this.groupBox1.Controls.Add(this.btnAnalyze);
            this.groupBox1.Controls.Add(this.txtUpKsRp);
            this.groupBox1.Controls.Add(this.txtExport);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnUpShRp);
            this.groupBox1.Controls.Add(this.txtUpShRp);
            this.groupBox1.Location = new System.Drawing.Point(10, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(348, 153);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Enabled = false;
            this.btnAnalyze.Location = new System.Drawing.Point(264, 89);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 5;
            this.btnAnalyze.Text = "处理";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // txtExport
            // 
            this.txtExport.Location = new System.Drawing.Point(74, 89);
            this.txtExport.Name = "txtExport";
            this.txtExport.Size = new System.Drawing.Size(183, 21);
            this.txtExport.TabIndex = 4;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(72, 128);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 128);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "操作消息:";
            // 
            // btnUpShRp
            // 
            this.btnUpShRp.Location = new System.Drawing.Point(263, 30);
            this.btnUpShRp.Name = "btnUpShRp";
            this.btnUpShRp.Size = new System.Drawing.Size(75, 23);
            this.btnUpShRp.TabIndex = 1;
            this.btnUpShRp.Text = "浏览";
            this.btnUpShRp.UseVisualStyleBackColor = true;
            this.btnUpShRp.Click += new System.EventHandler(this.btnUpShRp_Click);
            // 
            // txtUpShRp
            // 
            this.txtUpShRp.Enabled = false;
            this.txtUpShRp.Location = new System.Drawing.Point(74, 32);
            this.txtUpShRp.Name = "txtUpShRp";
            this.txtUpShRp.Size = new System.Drawing.Size(183, 21);
            this.txtUpShRp.TabIndex = 0;
            // 
            // txtUpKsRp
            // 
            this.txtUpKsRp.Enabled = false;
            this.txtUpKsRp.Location = new System.Drawing.Point(74, 62);
            this.txtUpKsRp.Name = "txtUpKsRp";
            this.txtUpKsRp.Size = new System.Drawing.Size(183, 21);
            this.txtUpKsRp.TabIndex = 4;
            // 
            // btnUpKsRp
            // 
            this.btnUpKsRp.Location = new System.Drawing.Point(264, 62);
            this.btnUpKsRp.Name = "btnUpKsRp";
            this.btnUpKsRp.Size = new System.Drawing.Size(75, 23);
            this.btnUpKsRp.TabIndex = 5;
            this.btnUpKsRp.Text = "浏览";
            this.btnUpKsRp.UseVisualStyleBackColor = true;
            this.btnUpKsRp.Click += new System.EventHandler(this.btnUpKsRp_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "上海仓库:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "昆山仓库:";
            // 
            // Form5
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(370, 162);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(386, 200);
            this.MinimumSize = new System.Drawing.Size(386, 200);
            this.Name = "Form5";
            this.Text = "延时报表";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.TextBox txtExport;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnUpShRp;
        private System.Windows.Forms.TextBox txtUpShRp;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnUpKsRp;
        private System.Windows.Forms.TextBox txtUpKsRp;
    }
}