namespace Gadget
{
    partial class _紧急单统计
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
            this.dtp截至时间 = new System.Windows.Forms.DateTimePicker();
            this.lbMsg = new System.Windows.Forms.Label();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.label3 = new System.Windows.Forms.Label();
            this.btn计算工作情况 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn上传紧急单 = new System.Windows.Forms.Button();
            this.txt紧急单 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dtp截至时间);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.btn计算工作情况);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btn上传紧急单);
            this.groupBox1.Controls.Add(this.txt紧急单);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(424, 108);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // dtp截至时间
            // 
            this.dtp截至时间.CustomFormat = "hh:mm";
            this.dtp截至时间.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtp截至时间.Location = new System.Drawing.Point(73, 48);
            this.dtp截至时间.Name = "dtp截至时间";
            this.dtp截至时间.ShowUpDown = true;
            this.dtp截至时间.Size = new System.Drawing.Size(108, 21);
            this.dtp截至时间.TabIndex = 14;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(71, 82);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 13;
            this.lbMsg.Text = "待上传文件";
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(365, 82);
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
            this.label3.Location = new System.Drawing.Point(6, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "操作消息:";
            // 
            // btn计算工作情况
            // 
            this.btn计算工作情况.Enabled = false;
            this.btn计算工作情况.Location = new System.Drawing.Point(341, 46);
            this.btn计算工作情况.Name = "btn计算工作情况";
            this.btn计算工作情况.Size = new System.Drawing.Size(75, 23);
            this.btn计算工作情况.TabIndex = 3;
            this.btn计算工作情况.Text = "计算";
            this.btn计算工作情况.UseVisualStyleBackColor = true;
            this.btn计算工作情况.Click += new System.EventHandler(this.btn计算工作情况_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "截至时间:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "紧急单:";
            // 
            // btn上传紧急单
            // 
            this.btn上传紧急单.Location = new System.Drawing.Point(341, 18);
            this.btn上传紧急单.Name = "btn上传紧急单";
            this.btn上传紧急单.Size = new System.Drawing.Size(75, 23);
            this.btn上传紧急单.TabIndex = 1;
            this.btn上传紧急单.Text = "浏览";
            this.btn上传紧急单.UseVisualStyleBackColor = true;
            this.btn上传紧急单.Click += new System.EventHandler(this.btn上传紧急单_Click);
            // 
            // txt紧急单
            // 
            this.txt紧急单.Enabled = false;
            this.txt紧急单.Location = new System.Drawing.Point(71, 20);
            this.txt紧急单.Name = "txt紧急单";
            this.txt紧急单.Size = new System.Drawing.Size(264, 21);
            this.txt紧急单.TabIndex = 0;
            // 
            // _紧急单统计
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(451, 131);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(467, 170);
            this.MinimumSize = new System.Drawing.Size(467, 170);
            this.Name = "_紧急单统计";
            this.Text = "_紧急单统计";
            this.Load += new System.EventHandler(this._紧急单统计_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn计算工作情况;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn上传紧急单;
        private System.Windows.Forms.TextBox txt紧急单;
        private System.Windows.Forms.DateTimePicker dtp截至时间;
    }
}