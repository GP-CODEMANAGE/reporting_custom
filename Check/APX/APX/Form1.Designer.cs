namespace APX
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
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.chkHistoricalTrxnPstn = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lblMessage = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnLoad = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.dGVSecurity = new System.Windows.Forms.DataGridView();
            this.txtHide = new System.Windows.Forms.TextBox();
            this.btnContinue = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cbUsedDate = new System.Windows.Forms.CheckBox();
            this.cbAccountSource = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lblSecurityMsg = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dGVSecurity)).BeginInit();
            this.SuspendLayout();
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(306, 87);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(15, 20);
            this.dateTimePicker1.TabIndex = 0;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // chkHistoricalTrxnPstn
            // 
            this.chkHistoricalTrxnPstn.AutoSize = true;
            this.chkHistoricalTrxnPstn.Location = new System.Drawing.Point(212, 165);
            this.chkHistoricalTrxnPstn.Name = "chkHistoricalTrxnPstn";
            this.chkHistoricalTrxnPstn.Size = new System.Drawing.Size(134, 17);
            this.chkHistoricalTrxnPstn.TabIndex = 22;
            this.chkHistoricalTrxnPstn.Text = "Append 9 to Port Code";
            this.chkHistoricalTrxnPstn.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(89, 199);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(78, 20);
            this.label3.TabIndex = 21;
            this.label3.Text = "Status :-";
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.lblMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMessage.Location = new System.Drawing.Point(208, 199);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(93, 20);
            this.lblMessage.TabIndex = 20;
            this.lblMessage.Text = "Data Load";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(194, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(152, 33);
            this.label1.TabIndex = 19;
            this.label1.Text = "APX Load";
            // 
            // btnLoad
            // 
            this.btnLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLoad.Location = new System.Drawing.Point(189, 272);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(175, 36);
            this.btnLoad.TabIndex = 18;
            this.btnLoad.Text = "Load Data";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(130, 87);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 13);
            this.label4.TabIndex = 23;
            this.label4.Text = "Select Date";
            // 
            // dGVSecurity
            // 
            this.dGVSecurity.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dGVSecurity.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dGVSecurity.Location = new System.Drawing.Point(12, 396);
            this.dGVSecurity.Name = "dGVSecurity";
            this.dGVSecurity.Size = new System.Drawing.Size(519, 182);
            this.dGVSecurity.TabIndex = 24;
            this.dGVSecurity.TabStop = false;
            this.dGVSecurity.Visible = false;
            // 
            // txtHide
            // 
            this.txtHide.Location = new System.Drawing.Point(212, 87);
            this.txtHide.Name = "txtHide";
            this.txtHide.Size = new System.Drawing.Size(89, 20);
            this.txtHide.TabIndex = 25;
            // 
            // btnContinue
            // 
            this.btnContinue.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnContinue.Location = new System.Drawing.Point(125, 327);
            this.btnContinue.Name = "btnContinue";
            this.btnContinue.Size = new System.Drawing.Size(118, 28);
            this.btnContinue.TabIndex = 26;
            this.btnContinue.Text = "Continue";
            this.btnContinue.UseVisualStyleBackColor = true;
            this.btnContinue.Visible = false;
            this.btnContinue.Click += new System.EventHandler(this.btnContinue_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(291, 327);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(118, 28);
            this.btnCancel.TabIndex = 27;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // cbUsedDate
            // 
            this.cbUsedDate.AutoSize = true;
            this.cbUsedDate.Location = new System.Drawing.Point(338, 90);
            this.cbUsedDate.Name = "cbUsedDate";
            this.cbUsedDate.Size = new System.Drawing.Size(71, 17);
            this.cbUsedDate.TabIndex = 28;
            this.cbUsedDate.Text = "Use Date";
            this.cbUsedDate.UseVisualStyleBackColor = true;
            this.cbUsedDate.CheckedChanged += new System.EventHandler(this.cbUsedDate_CheckedChanged);
            // 
            // cbAccountSource
            // 
            this.cbAccountSource.FormattingEnabled = true;
            this.cbAccountSource.Location = new System.Drawing.Point(212, 124);
            this.cbAccountSource.Name = "cbAccountSource";
            this.cbAccountSource.Size = new System.Drawing.Size(121, 21);
            this.cbAccountSource.TabIndex = 29;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(122, 132);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 30;
            this.label2.Text = "Account Source";
            // 
            // lblSecurityMsg
            // 
            this.lblSecurityMsg.AutoSize = true;
            this.lblSecurityMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSecurityMsg.Location = new System.Drawing.Point(43, 360);
            this.lblSecurityMsg.Name = "lblSecurityMsg";
            this.lblSecurityMsg.Size = new System.Drawing.Size(0, 13);
            this.lblSecurityMsg.TabIndex = 31;
            this.lblSecurityMsg.Visible = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(77, 240);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(441, 13);
            this.label5.TabIndex = 32;
            this.label5.Text = "* Please make sure that the posit.txt and transGA.txt files are closed before you" +
    " run the load*";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(543, 590);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lblSecurityMsg);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cbAccountSource);
            this.Controls.Add(this.cbUsedDate);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnContinue);
            this.Controls.Add(this.txtHide);
            this.Controls.Add(this.dGVSecurity);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.chkHistoricalTrxnPstn);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblMessage);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.dateTimePicker1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "APX Load";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dGVSecurity)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.CheckBox chkHistoricalTrxnPstn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridView dGVSecurity;
        private System.Windows.Forms.TextBox txtHide;
        private System.Windows.Forms.Button btnContinue;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckBox cbUsedDate;
        private System.Windows.Forms.ComboBox cbAccountSource;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblSecurityMsg;
        private System.Windows.Forms.Label label5;
    }
}

