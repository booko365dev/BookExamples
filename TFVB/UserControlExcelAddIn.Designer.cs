namespace TFVB
{
    partial class UserControlExcelAddIn
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnAddPicture = new System.Windows.Forms.Button();
            this.btnSaveAsCsv = new System.Windows.Forms.Button();
            this.btnGetTime = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnAddPicture
            // 
            this.btnAddPicture.Location = new System.Drawing.Point(25, 22);
            this.btnAddPicture.Name = "btnAddPicture";
            this.btnAddPicture.Size = new System.Drawing.Size(108, 57);
            this.btnAddPicture.TabIndex = 0;
            this.btnAddPicture.Text = "Add Picture";
            this.btnAddPicture.UseVisualStyleBackColor = true;
            this.btnAddPicture.Click += new System.EventHandler(this.BtnAddPicture_Click);
            // 
            // btnSaveAsCsv
            // 
            this.btnSaveAsCsv.Location = new System.Drawing.Point(25, 102);
            this.btnSaveAsCsv.Name = "btnSaveAsCsv";
            this.btnSaveAsCsv.Size = new System.Drawing.Size(108, 57);
            this.btnSaveAsCsv.TabIndex = 1;
            this.btnSaveAsCsv.Text = "Save As CSV";
            this.btnSaveAsCsv.UseVisualStyleBackColor = true;
            this.btnSaveAsCsv.Click += new System.EventHandler(this.BtnSaveAsCsv_Click);
            // 
            // btnGetTime
            // 
            this.btnGetTime.Location = new System.Drawing.Point(25, 186);
            this.btnGetTime.Name = "btnGetTime";
            this.btnGetTime.Size = new System.Drawing.Size(108, 57);
            this.btnGetTime.TabIndex = 2;
            this.btnGetTime.Text = "Get Time";
            this.btnGetTime.UseVisualStyleBackColor = true;
            this.btnGetTime.Click += new System.EventHandler(this.BtnGetTime_Click);
            // 
            // UserControlExcelAddIn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnGetTime);
            this.Controls.Add(this.btnSaveAsCsv);
            this.Controls.Add(this.btnAddPicture);
            this.Name = "UserControlExcelAddIn";
            this.Size = new System.Drawing.Size(150, 291);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnAddPicture;
        private System.Windows.Forms.Button btnSaveAsCsv;
        private System.Windows.Forms.Button btnGetTime;
    }
}
