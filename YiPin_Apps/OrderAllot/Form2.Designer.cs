namespace OrderAllot
{
    partial class Form2
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
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.txtExport = new System.Windows.Forms.TextBox();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUpload = new System.Windows.Forms.Button();
            this.txtUpload = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lkDecs);
            this.groupBox1.Controls.Add(this.btnAnalyze);
            this.groupBox1.Controls.Add(this.txtExport);
            this.groupBox1.Controls.Add(this.lbMsg);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnUpload);
            this.groupBox1.Controls.Add(this.txtUpload);
            this.groupBox1.Location = new System.Drawing.Point(11, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(334, 104);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据上传";
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(275, 82);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 23;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            this.lkDecs.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lkDecs_LinkClicked);
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Enabled = false;
            this.btnAnalyze.Location = new System.Drawing.Point(254, 48);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 5;
            this.btnAnalyze.Text = "处理";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // txtExport
            // 
            this.txtExport.Location = new System.Drawing.Point(64, 48);
            this.txtExport.Name = "txtExport";
            this.txtExport.Size = new System.Drawing.Size(183, 21);
            this.txtExport.TabIndex = 4;
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(72, 82);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 53);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "处理结果:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "工作情况:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "操作消息:";
            // 
            // btnUpload
            // 
            this.btnUpload.Location = new System.Drawing.Point(253, 18);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(75, 23);
            this.btnUpload.TabIndex = 1;
            this.btnUpload.Text = "浏览";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // txtUpload
            // 
            this.txtUpload.Enabled = false;
            this.txtUpload.Location = new System.Drawing.Point(64, 20);
            this.txtUpload.Name = "txtUpload";
            this.txtUpload.Size = new System.Drawing.Size(183, 21);
            this.txtUpload.TabIndex = 0;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(351, 116);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(367, 154);
            this.MinimumSize = new System.Drawing.Size(367, 154);
            this.Name = "Form2";
            this.Text = "工作完成情况统计";
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
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.TextBox txtUpload;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
    }
}