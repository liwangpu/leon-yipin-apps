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
            this.SuspendLayout();
            // 
            // btn浏览
            // 
            this.btn浏览.Location = new System.Drawing.Point(75, 89);
            this.btn浏览.Name = "btn浏览";
            this.btn浏览.Size = new System.Drawing.Size(75, 23);
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
            this.txtPath.Size = new System.Drawing.Size(218, 70);
            this.txtPath.TabIndex = 1;
            // 
            // btn合并
            // 
            this.btn合并.Location = new System.Drawing.Point(156, 89);
            this.btn合并.Name = "btn合并";
            this.btn合并.Size = new System.Drawing.Size(75, 23);
            this.btn合并.TabIndex = 2;
            this.btn合并.Text = "合并";
            this.btn合并.UseVisualStyleBackColor = true;
            this.btn合并.Click += new System.EventHandler(this.btn合并_Click);
            // 
            // _退款汇总表格合并
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(243, 121);
            this.Controls.Add(this.btn合并);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btn浏览);
            this.MaximumSize = new System.Drawing.Size(259, 159);
            this.MinimumSize = new System.Drawing.Size(259, 159);
            this.Name = "_退款汇总表格合并";
            this.Text = "_退款汇总表格合并";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn浏览;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Button btn合并;
    }
}