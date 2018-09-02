namespace Gadget
{
    partial class _日平均销量正态分布对比
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
            this.btn计算 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btn上传月销量流水 = new System.Windows.Forms.Button();
            this.txt月销量流水 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.btn计算);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btn上传月销量流水);
            this.groupBox1.Controls.Add(this.txt月销量流水);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(424, 115);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(69, 81);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 13;
            this.lbMsg.Text = "待上传文件";
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(363, 81);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 11;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 81);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "操作消息:";
            // 
            // btn计算
            // 
            this.btn计算.Enabled = false;
            this.btn计算.Location = new System.Drawing.Point(343, 47);
            this.btn计算.Name = "btn计算";
            this.btn计算.Size = new System.Drawing.Size(75, 23);
            this.btn计算.TabIndex = 3;
            this.btn计算.Text = "计算";
            this.btn计算.UseVisualStyleBackColor = true;
            this.btn计算.Click += new System.EventHandler(this.btn计算_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "销量详情:";
            // 
            // btn上传月销量流水
            // 
            this.btn上传月销量流水.Location = new System.Drawing.Point(341, 18);
            this.btn上传月销量流水.Name = "btn上传月销量流水";
            this.btn上传月销量流水.Size = new System.Drawing.Size(75, 23);
            this.btn上传月销量流水.TabIndex = 1;
            this.btn上传月销量流水.Text = "浏览";
            this.btn上传月销量流水.UseVisualStyleBackColor = true;
            this.btn上传月销量流水.Click += new System.EventHandler(this.btn上传月销量流水_Click);
            // 
            // txt月销量流水
            // 
            this.txt月销量流水.Enabled = false;
            this.txt月销量流水.Location = new System.Drawing.Point(71, 20);
            this.txt月销量流水.Name = "txt月销量流水";
            this.txt月销量流水.Size = new System.Drawing.Size(264, 21);
            this.txt月销量流水.TabIndex = 0;
            // 
            // _日平均销量正态分布对比
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(441, 134);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(457, 172);
            this.MinimumSize = new System.Drawing.Size(457, 172);
            this.Name = "_日平均销量正态分布对比";
            this.Text = "_日平均销量正态分布对比";
            this.Load += new System.EventHandler(this._日平均销量正态分布对比_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn计算;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn上传月销量流水;
        private System.Windows.Forms.TextBox txt月销量流水;
    }
}