namespace ManageMasterSchedule
{
    partial class SheetSpecForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SheetSpecForm));
            this.lblInstruction1 = new System.Windows.Forms.Label();
            this.btnSelectWord = new System.Windows.Forms.Button();
            this.lblWordDocLoc = new System.Windows.Forms.Label();
            this.txtWordPath = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProcessing1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // lblInstruction1
            // 
            this.lblInstruction1.AutoSize = true;
            this.lblInstruction1.Location = new System.Drawing.Point(13, 13);
            this.lblInstruction1.Name = "lblInstruction1";
            this.lblInstruction1.Size = new System.Drawing.Size(199, 13);
            this.lblInstruction1.TabIndex = 0;
            this.lblInstruction1.Text = "Select Master Schedule Excel document";
            // 
            // btnSelectWord
            // 
            this.btnSelectWord.Location = new System.Drawing.Point(16, 30);
            this.btnSelectWord.Name = "btnSelectWord";
            this.btnSelectWord.Size = new System.Drawing.Size(172, 23);
            this.btnSelectWord.TabIndex = 1;
            this.btnSelectWord.Text = "Select Excel Document";
            this.btnSelectWord.UseVisualStyleBackColor = true;
            this.btnSelectWord.Click += new System.EventHandler(this.btnSelectWord_Click);
            // 
            // lblWordDocLoc
            // 
            this.lblWordDocLoc.AutoSize = true;
            this.lblWordDocLoc.Location = new System.Drawing.Point(16, 60);
            this.lblWordDocLoc.Name = "lblWordDocLoc";
            this.lblWordDocLoc.Size = new System.Drawing.Size(123, 13);
            this.lblWordDocLoc.TabIndex = 2;
            this.lblWordDocLoc.Text = "Excel document location";
            // 
            // txtWordPath
            // 
            this.txtWordPath.Location = new System.Drawing.Point(19, 77);
            this.txtWordPath.Name = "txtWordPath";
            this.txtWordPath.Size = new System.Drawing.Size(313, 20);
            this.txtWordPath.TabIndex = 3;
            this.txtWordPath.TabStop = false;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(176, 135);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "Process";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(257, 135);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Close";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(19, 134);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(151, 23);
            this.progressBar1.TabIndex = 4;
            this.progressBar1.Visible = false;
            // 
            // lblProcessing1
            // 
            this.lblProcessing1.AutoSize = true;
            this.lblProcessing1.Location = new System.Drawing.Point(16, 117);
            this.lblProcessing1.Name = "lblProcessing1";
            this.lblProcessing1.Size = new System.Drawing.Size(35, 13);
            this.lblProcessing1.TabIndex = 5;
            this.lblProcessing1.Text = "label1";
            this.lblProcessing1.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ManageMasterSchedule.Properties.Resources.processing_32;
            this.pictureBox1.Location = new System.Drawing.Point(300, 100);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(32, 32);
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            // 
            // SheetSpecForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(344, 172);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lblProcessing1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtWordPath);
            this.Controls.Add(this.lblWordDocLoc);
            this.Controls.Add(this.btnSelectWord);
            this.Controls.Add(this.lblInstruction1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SheetSpecForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Manage Master Schedule";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblInstruction1;
        private System.Windows.Forms.Button btnSelectWord;
        private System.Windows.Forms.Label lblWordDocLoc;
        private System.Windows.Forms.TextBox txtWordPath;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProcessing1;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}