namespace Gadget
{
    partial class _批量生成采购订单
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
            this.lbMsg = new System.Windows.Forms.Label();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.label3 = new System.Windows.Forms.Label();
            this.btn处理数据 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn上传库存预警原表 = new System.Windows.Forms.Button();
            this.txt库存预警原表 = new System.Windows.Forms.TextBox();
            this.btn上传库存预警中位数 = new System.Windows.Forms.Button();
            this.btn上传每月流水 = new System.Windows.Forms.Button();
            this.txt库存预警中位数 = new System.Windows.Forms.TextBox();
            this.txt每月流水 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.btn处理数据);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btn上传每月流水);
            this.groupBox1.Controls.Add(this.btn上传库存预警中位数);
            this.groupBox1.Controls.Add(this.btn上传库存预警原表);
            this.groupBox1.Controls.Add(this.txt每月流水);
            this.groupBox1.Controls.Add(this.txt库存预警中位数);
            this.groupBox1.Controls.Add(this.txt库存预警原表);
            this.groupBox1.Location = new System.Drawing.Point(12, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(468, 169);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(71, 144);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 13;
            this.lbMsg.Text = "待上传文件";
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(404, 144);
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
            this.label3.Location = new System.Drawing.Point(6, 144);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "操作消息:";
            // 
            // btn处理数据
            // 
            this.btn处理数据.Enabled = false;
            this.btn处理数据.Location = new System.Drawing.Point(382, 109);
            this.btn处理数据.Name = "btn处理数据";
            this.btn处理数据.Size = new System.Drawing.Size(75, 23);
            this.btn处理数据.TabIndex = 3;
            this.btn处理数据.Text = "计算";
            this.btn处理数据.UseVisualStyleBackColor = true;
            this.btn处理数据.Click += new System.EventHandler(this.btn处理数据_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(101, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "库存预警-中位数:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "库存预警-原先:";
            // 
            // btn上传库存预警原表
            // 
            this.btn上传库存预警原表.Location = new System.Drawing.Point(382, 18);
            this.btn上传库存预警原表.Name = "btn上传库存预警原表";
            this.btn上传库存预警原表.Size = new System.Drawing.Size(75, 23);
            this.btn上传库存预警原表.TabIndex = 1;
            this.btn上传库存预警原表.Text = "浏览";
            this.btn上传库存预警原表.UseVisualStyleBackColor = true;
            this.btn上传库存预警原表.Click += new System.EventHandler(this.btn上传库存预警原表_Click);
            // 
            // txt库存预警原表
            // 
            this.txt库存预警原表.Enabled = false;
            this.txt库存预警原表.Location = new System.Drawing.Point(112, 20);
            this.txt库存预警原表.Name = "txt库存预警原表";
            this.txt库存预警原表.Size = new System.Drawing.Size(264, 21);
            this.txt库存预警原表.TabIndex = 0;
            // 
            // btn上传库存预警中位数
            // 
            this.btn上传库存预警中位数.Location = new System.Drawing.Point(382, 49);
            this.btn上传库存预警中位数.Name = "btn上传库存预警中位数";
            this.btn上传库存预警中位数.Size = new System.Drawing.Size(75, 23);
            this.btn上传库存预警中位数.TabIndex = 1;
            this.btn上传库存预警中位数.Text = "浏览";
            this.btn上传库存预警中位数.UseVisualStyleBackColor = true;
            this.btn上传库存预警中位数.Click += new System.EventHandler(this.btn上传库存预警中位数_Click);
            // 
            // btn上传每月流水
            // 
            this.btn上传每月流水.Location = new System.Drawing.Point(382, 80);
            this.btn上传每月流水.Name = "btn上传每月流水";
            this.btn上传每月流水.Size = new System.Drawing.Size(75, 23);
            this.btn上传每月流水.TabIndex = 1;
            this.btn上传每月流水.Text = "浏览";
            this.btn上传每月流水.UseVisualStyleBackColor = true;
            this.btn上传每月流水.Click += new System.EventHandler(this.btn上传每月流水_Click);
            // 
            // txt库存预警中位数
            // 
            this.txt库存预警中位数.Enabled = false;
            this.txt库存预警中位数.Location = new System.Drawing.Point(114, 51);
            this.txt库存预警中位数.Name = "txt库存预警中位数";
            this.txt库存预警中位数.Size = new System.Drawing.Size(264, 21);
            this.txt库存预警中位数.TabIndex = 0;
            // 
            // txt每月流水
            // 
            this.txt每月流水.Enabled = false;
            this.txt每月流水.Location = new System.Drawing.Point(114, 80);
            this.txt每月流水.Name = "txt每月流水";
            this.txt每月流水.Size = new System.Drawing.Size(264, 21);
            this.txt每月流水.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "每月销售流水:";
            // 
            // _批量生成采购订单
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(491, 176);
            this.Controls.Add(this.groupBox1);
            this.Name = "_批量生成采购订单";
            this.Text = "_批量生成采购订单";
            this.Load += new System.EventHandler(this._批量生成采购订单_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn处理数据;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn上传每月流水;
        private System.Windows.Forms.Button btn上传库存预警中位数;
        private System.Windows.Forms.Button btn上传库存预警原表;
        private System.Windows.Forms.TextBox txt每月流水;
        private System.Windows.Forms.TextBox txt库存预警中位数;
        private System.Windows.Forms.TextBox txt库存预警原表;
    }
}