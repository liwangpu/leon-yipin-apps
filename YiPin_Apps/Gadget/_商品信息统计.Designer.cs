﻿namespace Gadget
{
    partial class _商品信息统计
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
            this.ndLower = new System.Windows.Forms.NumericUpDown();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.txtExport = new System.Windows.Forms.TextBox();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUploadBuyer = new System.Windows.Forms.Button();
            this.btn商品明细 = new System.Windows.Forms.Button();
            this.txtBuyer = new System.Windows.Forms.TextBox();
            this.txt商品明细 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ndLower)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ndLower);
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btnAnalyze);
            this.groupBox1.Controls.Add(this.txtExport);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnUploadBuyer);
            this.groupBox1.Controls.Add(this.btn商品明细);
            this.groupBox1.Controls.Add(this.txtBuyer);
            this.groupBox1.Controls.Add(this.txt商品明细);
            this.groupBox1.Location = new System.Drawing.Point(12, 7);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(357, 165);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // ndLower
            // 
            this.ndLower.Location = new System.Drawing.Point(75, 113);
            this.ndLower.Maximum = new decimal(new int[] {
            1410065407,
            2,
            0,
            0});
            this.ndLower.Name = "ndLower";
            this.ndLower.Size = new System.Drawing.Size(183, 21);
            this.ndLower.TabIndex = 21;
            this.ndLower.Value = new decimal(new int[] {
            5000,
            0,
            0,
            0});
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(286, 137);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 20;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            this.lkDecs.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkDecs_LinkClicked);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(155, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "供应商详情月销量金额下限:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(31, 77);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(35, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "结果:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(19, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "采购员:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "商品明细:";
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Location = new System.Drawing.Point(264, 72);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 5;
            this.btnAnalyze.Text = "处理";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // txtExport
            // 
            this.txtExport.Location = new System.Drawing.Point(75, 74);
            this.txtExport.Name = "txtExport";
            this.txtExport.Size = new System.Drawing.Size(183, 21);
            this.txtExport.TabIndex = 4;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(73, 137);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 137);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "操作消息:";
            // 
            // btnUploadBuyer
            // 
            this.btnUploadBuyer.Location = new System.Drawing.Point(264, 45);
            this.btnUploadBuyer.Name = "btnUploadBuyer";
            this.btnUploadBuyer.Size = new System.Drawing.Size(75, 23);
            this.btnUploadBuyer.TabIndex = 1;
            this.btnUploadBuyer.Text = "浏览";
            this.btnUploadBuyer.UseVisualStyleBackColor = true;
            this.btnUploadBuyer.Click += new System.EventHandler(this.btnUploadBuyer_Click);
            // 
            // btn商品明细
            // 
            this.btn商品明细.Location = new System.Drawing.Point(264, 18);
            this.btn商品明细.Name = "btn商品明细";
            this.btn商品明细.Size = new System.Drawing.Size(75, 23);
            this.btn商品明细.TabIndex = 1;
            this.btn商品明细.Text = "浏览";
            this.btn商品明细.UseVisualStyleBackColor = true;
            this.btn商品明细.Click += new System.EventHandler(this.btn商品明细_Click);
            // 
            // txtBuyer
            // 
            this.txtBuyer.Enabled = false;
            this.txtBuyer.Location = new System.Drawing.Point(75, 47);
            this.txtBuyer.Name = "txtBuyer";
            this.txtBuyer.Size = new System.Drawing.Size(183, 21);
            this.txtBuyer.TabIndex = 0;
            // 
            // txt商品明细
            // 
            this.txt商品明细.Enabled = false;
            this.txt商品明细.Location = new System.Drawing.Point(75, 20);
            this.txt商品明细.Name = "txt商品明细";
            this.txt商品明细.Size = new System.Drawing.Size(183, 21);
            this.txt商品明细.TabIndex = 0;
            // 
            // _商品信息统计
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(375, 177);
            this.Controls.Add(this.groupBox1);
            this.MaximumSize = new System.Drawing.Size(391, 216);
            this.MinimumSize = new System.Drawing.Size(391, 216);
            this.Name = "_商品信息统计";
            this.Text = "_商品信息统计";
            this.Load += new System.EventHandler(this._商品信息统计_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ndLower)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.TextBox txtExport;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn商品明细;
        private System.Windows.Forms.TextBox txt商品明细;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown ndLower;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnUploadBuyer;
        private System.Windows.Forms.TextBox txtBuyer;
    }
}