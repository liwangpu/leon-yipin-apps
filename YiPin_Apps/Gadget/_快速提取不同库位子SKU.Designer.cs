namespace Gadget
{
    partial class _快速提取不同库位子SKU
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnCalcu = new System.Windows.Forms.Button();
            this.lkDecs = new System.Windows.Forms.LinkLabel();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtFloor = new System.Windows.Forms.TextBox();
            this.btnUploadFloor = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "SKU:";
            // 
            // txtFile
            // 
            this.txtFile.Location = new System.Drawing.Point(67, 17);
            this.txtFile.Name = "txtFile";
            this.txtFile.ReadOnly = true;
            this.txtFile.Size = new System.Drawing.Size(219, 21);
            this.txtFile.TabIndex = 1;
            // 
            // btnUpload
            // 
            this.btnUpload.Location = new System.Drawing.Point(292, 15);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(64, 23);
            this.btnUpload.TabIndex = 2;
            this.btnUpload.Text = "上传";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.BtnUpload_Click);
            // 
            // btnCalcu
            // 
            this.btnCalcu.Location = new System.Drawing.Point(292, 71);
            this.btnCalcu.Name = "btnCalcu";
            this.btnCalcu.Size = new System.Drawing.Size(64, 23);
            this.btnCalcu.TabIndex = 3;
            this.btnCalcu.Text = "处理";
            this.btnCalcu.UseVisualStyleBackColor = true;
            this.btnCalcu.Click += new System.EventHandler(this.BtnCalcu_Click);
            // 
            // lkDecs
            // 
            this.lkDecs.AutoSize = true;
            this.lkDecs.Location = new System.Drawing.Point(10, 98);
            this.lkDecs.Name = "lkDecs";
            this.lkDecs.Size = new System.Drawing.Size(53, 12);
            this.lkDecs.TabIndex = 21;
            this.lkDecs.TabStop = true;
            this.lkDecs.Text = "表格说明";
            this.lkDecs.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LkDecs_LinkClicked_1);
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(75, 76);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 23;
            this.lbMsg.Text = "待上传文件";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 76);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 12);
            this.label2.TabIndex = 22;
            this.label2.Text = "操作消息:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "楼层:";
            // 
            // txtFloor
            // 
            this.txtFloor.Location = new System.Drawing.Point(67, 42);
            this.txtFloor.Name = "txtFloor";
            this.txtFloor.ReadOnly = true;
            this.txtFloor.Size = new System.Drawing.Size(219, 21);
            this.txtFloor.TabIndex = 1;
            // 
            // btnUploadFloor
            // 
            this.btnUploadFloor.Location = new System.Drawing.Point(292, 42);
            this.btnUploadFloor.Name = "btnUploadFloor";
            this.btnUploadFloor.Size = new System.Drawing.Size(64, 23);
            this.btnUploadFloor.TabIndex = 2;
            this.btnUploadFloor.Text = "上传";
            this.btnUploadFloor.UseVisualStyleBackColor = true;
            this.btnUploadFloor.Click += new System.EventHandler(this.BtnUploadFloor_Click);
            // 
            // _快速提取不同库位子SKU
            // 
            this.ClientSize = new System.Drawing.Size(368, 120);
            this.Controls.Add(this.lbMsg);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lkDecs);
            this.Controls.Add(this.btnCalcu);
            this.Controls.Add(this.btnUploadFloor);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.txtFloor);
            this.Controls.Add(this.txtFile);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.Name = "_快速提取不同库位子SKU";
            this.Text = "快速提取不同库位子SKU";
            this.Load += new System.EventHandler(this._快速提取不同库位子SKU_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnCalcu;
        private System.Windows.Forms.LinkLabel lkDecs;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtFloor;
        private System.Windows.Forms.Button btnUploadFloor;
    }
}