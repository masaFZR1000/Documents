namespace EMS_Small_Chb
{
    partial class SetCheck
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
            this.formsPlot1 = new ScottPlot.FormsPlot();
            this.lblTestCheckMsg1 = new System.Windows.Forms.Label();
            this.lblTestCheckMsg2 = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // formsPlot1
            // 
            this.formsPlot1.Location = new System.Drawing.Point(12, 345);
            this.formsPlot1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.formsPlot1.Name = "formsPlot1";
            this.formsPlot1.Size = new System.Drawing.Size(1240, 565);
            this.formsPlot1.TabIndex = 0;
            // 
            // lblTestCheckMsg1
            // 
            this.lblTestCheckMsg1.Font = new System.Drawing.Font("MS UI Gothic", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblTestCheckMsg1.Location = new System.Drawing.Point(10, 9);
            this.lblTestCheckMsg1.Name = "lblTestCheckMsg1";
            this.lblTestCheckMsg1.Size = new System.Drawing.Size(600, 334);
            this.lblTestCheckMsg1.TabIndex = 1;
            this.lblTestCheckMsg1.Text = "label1";
            // 
            // lblTestCheckMsg2
            // 
            this.lblTestCheckMsg2.Font = new System.Drawing.Font("MS UI Gothic", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblTestCheckMsg2.Location = new System.Drawing.Point(616, 9);
            this.lblTestCheckMsg2.Name = "lblTestCheckMsg2";
            this.lblTestCheckMsg2.Size = new System.Drawing.Size(498, 334);
            this.lblTestCheckMsg2.TabIndex = 2;
            this.lblTestCheckMsg2.Text = "label1";
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnClose.Location = new System.Drawing.Point(1113, 24);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(121, 321);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "閉じる";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // SetCheck
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1264, 921);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.lblTestCheckMsg2);
            this.Controls.Add(this.lblTestCheckMsg1);
            this.Controls.Add(this.formsPlot1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SetCheck";
            this.Text = "SetCheck";
            this.ResumeLayout(false);

        }

        #endregion

        private ScottPlot.FormsPlot formsPlot1;
        private System.Windows.Forms.Label lblTestCheckMsg1;
        private System.Windows.Forms.Label lblTestCheckMsg2;
        private System.Windows.Forms.Button btnClose;
    }
}