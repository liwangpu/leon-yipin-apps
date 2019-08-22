namespace Gadget
{
    partial class _库位号订单提取
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
            this.label2 = new System.Windows.Forms.Label();
            this.btnCalcu = new System.Windows.Forms.Button();
            this.btn上传订单表 = new System.Windows.Forms.Button();
            this.txt订单表 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txt库位号 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnCalcu);
            this.groupBox1.Controls.Add(this.btn上传订单表);
            this.groupBox1.Controls.Add(this.txt库位号);
            this.groupBox1.Controls.Add(this.txt订单表);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(424, 102);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(69, 79);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 13;
            this.lbMsg.Text = "待上传文件";
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(363, 79);
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
            this.label3.Location = new System.Drawing.Point(6, 79);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "操作消息:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "订单表:";
            // 
            // btnCalcu
            // 
            this.btnCalcu.Location = new System.Drawing.Point(341, 47);
            this.btnCalcu.Name = "btnCalcu";
            this.btnCalcu.Size = new System.Drawing.Size(75, 23);
            this.btnCalcu.TabIndex = 1;
            this.btnCalcu.Text = "计算";
            this.btnCalcu.UseVisualStyleBackColor = true;
            this.btnCalcu.Click += new System.EventHandler(this.btnCalcu_Click);
            // 
            // btn上传订单表
            // 
            this.btn上传订单表.Location = new System.Drawing.Point(341, 18);
            this.btn上传订单表.Name = "btn上传订单表";
            this.btn上传订单表.Size = new System.Drawing.Size(75, 23);
            this.btn上传订单表.TabIndex = 1;
            this.btn上传订单表.Text = "浏览";
            this.btn上传订单表.UseVisualStyleBackColor = true;
            this.btn上传订单表.Click += new System.EventHandler(this.btn上传订单表_Click);
            // 
            // txt订单表
            // 
            this.txt订单表.Enabled = false;
            this.txt订单表.Location = new System.Drawing.Point(71, 20);
            this.txt订单表.Name = "txt订单表";
            this.txt订单表.Size = new System.Drawing.Size(264, 21);
            this.txt订单表.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "库位号:";
            // 
            // txt库位号
            // 
            this.txt库位号.Location = new System.Drawing.Point(71, 47);
            this.txt库位号.Name = "txt库位号";
            this.txt库位号.Size = new System.Drawing.Size(264, 21);
            this.txt库位号.TabIndex = 0;
            this.txt库位号.Text = "A";
            // 
            // _库位号订单提取
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(441, 119);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(457, 158);
            this.MinimumSize = new System.Drawing.Size(457, 158);
            this.Name = "_库位号订单提取";
            this.Text = "_库位号订单提取";
            this.Load += new System.EventHandler(this._库位号订单提取_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn上传订单表;
        private System.Windows.Forms.TextBox txt订单表;
        private System.Windows.Forms.Button btnCalcu;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt库位号;
    }
}