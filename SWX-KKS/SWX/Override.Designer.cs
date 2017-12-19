namespace SWX_KKS.SWX
{
    partial class Override
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
            this.btnCancel = new System.Windows.Forms.Button();
            this.cbRemember = new System.Windows.Forms.CheckBox();
            this.lblText = new System.Windows.Forms.Label();
            this.btnOverride = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Image = global::SWX_KKS.Properties.Resources.KKS_Button;
            this.btnCancel.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btnCancel.Location = new System.Drawing.Point(313, 141);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(84, 26);
            this.btnCancel.TabIndex = 48;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // cbRemember
            // 
            this.cbRemember.AutoSize = true;
            this.cbRemember.Location = new System.Drawing.Point(12, 87);
            this.cbRemember.Name = "cbRemember";
            this.cbRemember.Size = new System.Drawing.Size(103, 17);
            this.cbRemember.TabIndex = 49;
            this.cbRemember.Text = "Eingabe merken";
            this.cbRemember.UseVisualStyleBackColor = true;
            // 
            // lblText
            // 
            this.lblText.AutoSize = true;
            this.lblText.Location = new System.Drawing.Point(12, 62);
            this.lblText.Name = "lblText";
            this.lblText.Size = new System.Drawing.Size(35, 13);
            this.lblText.TabIndex = 50;
            this.lblText.Text = "label1";
            // 
            // btnOverride
            // 
            this.btnOverride.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOverride.ForeColor = System.Drawing.Color.White;
            this.btnOverride.Image = global::SWX_KKS.Properties.Resources.KKS_Button;
            this.btnOverride.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
            this.btnOverride.Location = new System.Drawing.Point(223, 141);
            this.btnOverride.Name = "btnOverride";
            this.btnOverride.Size = new System.Drawing.Size(84, 26);
            this.btnOverride.TabIndex = 48;
            this.btnOverride.Text = "Ja";
            this.btnOverride.UseVisualStyleBackColor = true;
            this.btnOverride.Click += new System.EventHandler(this.btnOverride_Click);
            // 
            // Override
            // 
            this.AcceptButton = this.btnOverride;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(409, 228);
            this.Controls.Add(this.lblText);
            this.Controls.Add(this.cbRemember);
            this.Controls.Add(this.btnOverride);
            this.Controls.Add(this.btnCancel);
            this.Name = "Override";
            this.Text = "Override";
            this.Load += new System.EventHandler(this.Override_Load);
            this.Controls.SetChildIndex(this.btnCancel, 0);
            this.Controls.SetChildIndex(this.btnOverride, 0);
            this.Controls.SetChildIndex(this.cbRemember, 0);
            this.Controls.SetChildIndex(this.lblText, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckBox cbRemember;
        private System.Windows.Forms.Label lblText;
        private System.Windows.Forms.Button btnOverride;
    }
}