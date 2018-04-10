namespace SummaryGenerator
{
    partial class Form1
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
            this.btnAdd_Incidents = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.rbMs4 = new System.Windows.Forms.RadioButton();
            this.rbInitiate = new System.Windows.Forms.RadioButton();
            this.txtSrcFilePath = new System.Windows.Forms.TextBox();
            this.txtDestFilePath = new System.Windows.Forms.TextBox();
            this.btnSrcFileDialog = new System.Windows.Forms.Button();
            this.btnDestFileDialog = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnExit = new System.Windows.Forms.Button();
            this.txtSecName = new System.Windows.Forms.TextBox();
            this.lblSecName = new System.Windows.Forms.Label();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnAdd_Incidents
            // 
            this.btnAdd_Incidents.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnAdd_Incidents.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAdd_Incidents.Location = new System.Drawing.Point(52, 396);
            this.btnAdd_Incidents.Name = "btnAdd_Incidents";
            this.btnAdd_Incidents.Size = new System.Drawing.Size(276, 57);
            this.btnAdd_Incidents.TabIndex = 0;
            this.btnAdd_Incidents.Text = "Update Incident Tracker";
            this.btnAdd_Incidents.UseVisualStyleBackColor = false;
            this.btnAdd_Incidents.Click += new System.EventHandler(this.btnAdd_Incidents_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(143, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(216, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "  Updating Incidents for:  ";
            // 
            // rbMs4
            // 
            this.rbMs4.AutoSize = true;
            this.rbMs4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rbMs4.Location = new System.Drawing.Point(170, 66);
            this.rbMs4.Name = "rbMs4";
            this.rbMs4.Size = new System.Drawing.Size(64, 24);
            this.rbMs4.TabIndex = 2;
            this.rbMs4.TabStop = true;
            this.rbMs4.Text = "MS4";
            this.rbMs4.UseVisualStyleBackColor = true;
            this.rbMs4.CheckedChanged += new System.EventHandler(this.rbMs4_CheckedChanged);
            // 
            // rbInitiate
            // 
            this.rbInitiate.AutoSize = true;
            this.rbInitiate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rbInitiate.Location = new System.Drawing.Point(249, 66);
            this.rbInitiate.Name = "rbInitiate";
            this.rbInitiate.Size = new System.Drawing.Size(79, 24);
            this.rbInitiate.TabIndex = 3;
            this.rbInitiate.TabStop = true;
            this.rbInitiate.Text = "Initiate";
            this.rbInitiate.UseVisualStyleBackColor = true;
            this.rbInitiate.CheckedChanged += new System.EventHandler(this.rbInitiate_CheckedChanged);
            // 
            // txtSrcFilePath
            // 
            this.txtSrcFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSrcFilePath.Location = new System.Drawing.Point(45, 233);
            this.txtSrcFilePath.Name = "txtSrcFilePath";
            this.txtSrcFilePath.Size = new System.Drawing.Size(349, 24);
            this.txtSrcFilePath.TabIndex = 4;
            // 
            // txtDestFilePath
            // 
            this.txtDestFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDestFilePath.Location = new System.Drawing.Point(45, 306);
            this.txtDestFilePath.Name = "txtDestFilePath";
            this.txtDestFilePath.Size = new System.Drawing.Size(349, 24);
            this.txtDestFilePath.TabIndex = 5;
            // 
            // btnSrcFileDialog
            // 
            this.btnSrcFileDialog.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSrcFileDialog.Location = new System.Drawing.Point(401, 228);
            this.btnSrcFileDialog.Name = "btnSrcFileDialog";
            this.btnSrcFileDialog.Size = new System.Drawing.Size(75, 34);
            this.btnSrcFileDialog.TabIndex = 6;
            this.btnSrcFileDialog.Text = "Browse";
            this.btnSrcFileDialog.UseVisualStyleBackColor = true;
            this.btnSrcFileDialog.Click += new System.EventHandler(this.btnSrcFileDialog_Click);
            // 
            // btnDestFileDialog
            // 
            this.btnDestFileDialog.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDestFileDialog.Location = new System.Drawing.Point(400, 302);
            this.btnDestFileDialog.Name = "btnDestFileDialog";
            this.btnDestFileDialog.Size = new System.Drawing.Size(75, 32);
            this.btnDestFileDialog.TabIndex = 7;
            this.btnDestFileDialog.Text = "Browse";
            this.btnDestFileDialog.UseVisualStyleBackColor = true;
            this.btnDestFileDialog.Click += new System.EventHandler(this.btnDestFileDialog_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(41, 212);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(275, 18);
            this.label2.TabIndex = 8;
            this.label2.Text = "Source file path (From Service Matters) :";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(42, 283);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(263, 18);
            this.label3.TabIndex = 9;
            this.label3.Text = "Destination file path (Incident Tracker) :";
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = System.Drawing.Color.PaleGreen;
            this.statusStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible;
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 504);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(526, 25);
            this.statusStrip1.TabIndex = 10;
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(151, 20);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // btnExit
            // 
            this.btnExit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExit.Location = new System.Drawing.Point(338, 396);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(106, 57);
            this.btnExit.TabIndex = 11;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // txtSecName
            // 
            this.txtSecName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSecName.Location = new System.Drawing.Point(45, 156);
            this.txtSecName.Name = "txtSecName";
            this.txtSecName.Size = new System.Drawing.Size(283, 24);
            this.txtSecName.TabIndex = 12;
            // 
            // lblSecName
            // 
            this.lblSecName.AutoSize = true;
            this.lblSecName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSecName.Location = new System.Drawing.Point(42, 135);
            this.lblSecName.Name = "lblSecName";
            this.lblSecName.Size = new System.Drawing.Size(175, 18);
            this.lblSecName.TabIndex = 13;
            this.lblSecName.Text = "Secondary Person Name";
            // 
            // Form1
            // 
            this.AcceptButton = this.btnAdd_Incidents;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 529);
            this.Controls.Add(this.lblSecName);
            this.Controls.Add(this.txtSecName);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnDestFileDialog);
            this.Controls.Add(this.btnSrcFileDialog);
            this.Controls.Add(this.txtDestFilePath);
            this.Controls.Add(this.txtSrcFilePath);
            this.Controls.Add(this.rbInitiate);
            this.Controls.Add(this.rbMs4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAdd_Incidents);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Incident Tracker";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAdd_Incidents;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton rbMs4;
        private System.Windows.Forms.RadioButton rbInitiate;
        private System.Windows.Forms.TextBox txtSrcFilePath;
        private System.Windows.Forms.TextBox txtDestFilePath;
        private System.Windows.Forms.Button btnSrcFileDialog;
        private System.Windows.Forms.Button btnDestFileDialog;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.TextBox txtSecName;
        private System.Windows.Forms.Label lblSecName;
    }
}

