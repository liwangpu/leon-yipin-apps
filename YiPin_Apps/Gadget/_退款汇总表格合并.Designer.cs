namespace Gadget
{
    partial class _退款汇总表格合并
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
            this.btn浏览 = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.btn合并 = new System.Windows.Forms.Button();
            this.lbMsg = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btn交易数量汇总 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn浏览
            // 
            this.btn浏览.Location = new System.Drawing.Point(12, 89);
            this.btn浏览.Name = "btn浏览";
            this.btn浏览.Size = new System.Drawing.Size(106, 23);
            this.btn浏览.TabIndex = 0;
            this.btn浏览.Text = "浏览";
            this.btn浏览.UseVisualStyleBackColor = true;
            this.btn浏览.Click += new System.EventHandler(this.btn浏览_Click);
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(12, 12);
            this.txtPath.Multiline = true;
            this.txtPath.Name = "txtPath";
            this.txtPath.ReadOnly = true;
            this.txtPath.Size = new System.Drawing.Size(218, 70);
            this.txtPath.TabIndex = 1;
            // 
            // btn合并
            // 
            this.btn合并.Location = new System.Drawing.Point(124, 89);
            this.btn合并.Name = "btn合并";
            this.btn合并.Size = new System.Drawing.Size(107, 23);
            this.btn合并.TabIndex = 2;
            this.btn合并.Text = "合并汇总";
            this.btn合并.UseVisualStyleBackColor = true;
            this.btn合并.Click += new System.EventHandler(this.btn合并_Click);
            // 
            // lbMsg
            // 
            this.lbMsg.AutoSize = true;
            this.lbMsg.Location = new System.Drawing.Point(53, 153);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(65, 12);
            this.lbMsg.TabIndex = 3;
            this.lbMsg.Text = "待上传文件";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 153);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "状态:";
            // 
            // btn交易数量汇总
            // 
            this.btn交易数量汇总.Location = new System.Drawing.Point(124, 119);
            this.btn交易数量汇总.Name = "btn交易数量汇总";
            this.btn交易数量汇总.Size = new System.Drawing.Size(107, 23);
            this.btn交易数量汇总.TabIndex = 4;
            this.btn交易数量汇总.Text = "交易数量汇总";
            this.btn交易数量汇总.UseVisualStyleBackColor = true;
            this.btn交易数量汇总.Click += new System.EventHandler(this.btn交易数量汇总_Click);
            // 
            // _退款汇总表格合并
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(249, 174);
            this.Controls.Add(this.btn交易数量汇总);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbMsg);
            this.Controls.Add(this.btn合并);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btn浏览);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(265, 213);
            this.MinimumSize = new System.Drawing.Size(265, 213);
            this.Name = "_退款汇总表格合并";
            this.Text = "退款汇总表格合并";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn浏览;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btn合并;
        private System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn交易数量汇总;
    }
}