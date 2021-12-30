namespace GPTool
{
    partial class GPTool
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GPTool));
            this.excelPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.billNum = new System.Windows.Forms.TextBox();
            this.ErrorMsg = new System.Windows.Forms.Label();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.printPreviewDialog1 = new System.Windows.Forms.PrintPreviewDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.printRange = new System.Windows.Forms.Button();
            this.billNoFrom = new System.Windows.Forms.TextBox();
            this.billNoTo = new System.Windows.Forms.TextBox();
            this.printAll = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // excelPath
            // 
            this.excelPath.Location = new System.Drawing.Point(284, 95);
            this.excelPath.Name = "excelPath";
            this.excelPath.Size = new System.Drawing.Size(264, 26);
            this.excelPath.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(186, 101);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "Excel Path";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(578, 140);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(95, 39);
            this.button1.TabIndex = 2;
            this.button1.Text = "Print";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(186, 156);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 20);
            this.label2.TabIndex = 3;
            this.label2.Text = "Bill No:";
            // 
            // billNum
            // 
            this.billNum.Location = new System.Drawing.Point(284, 153);
            this.billNum.Name = "billNum";
            this.billNum.Size = new System.Drawing.Size(264, 26);
            this.billNum.TabIndex = 4;
            this.billNum.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.billNum_KeyPress);
            // 
            // ErrorMsg
            // 
            this.ErrorMsg.AutoSize = true;
            this.ErrorMsg.Location = new System.Drawing.Point(280, 265);
            this.ErrorMsg.Name = "ErrorMsg";
            this.ErrorMsg.Size = new System.Drawing.Size(137, 20);
            this.ErrorMsg.TabIndex = 5;
            this.ErrorMsg.Text = "Something Wrong";
            // 
            // printDocument1
            // 
            this.printDocument1.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument1_PrintPage);
            // 
            // printPreviewDialog1
            // 
            this.printPreviewDialog1.AutoScrollMargin = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.AutoScrollMinSize = new System.Drawing.Size(0, 0);
            this.printPreviewDialog1.ClientSize = new System.Drawing.Size(400, 300);
            this.printPreviewDialog1.Document = this.printDocument1;
            this.printPreviewDialog1.Enabled = true;
            this.printPreviewDialog1.Icon = ((System.Drawing.Icon)(resources.GetObject("printPreviewDialog1.Icon")));
            this.printPreviewDialog1.Name = "printPreviewDialog1";
            this.printPreviewDialog1.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(186, 221);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 20);
            this.label3.TabIndex = 6;
            this.label3.Text = "Bill No:";
            // 
            // printRange
            // 
            this.printRange.Location = new System.Drawing.Point(578, 199);
            this.printRange.Name = "printRange";
            this.printRange.Size = new System.Drawing.Size(95, 39);
            this.printRange.TabIndex = 7;
            this.printRange.Text = "Print";
            this.printRange.UseVisualStyleBackColor = true;
            this.printRange.Click += new System.EventHandler(this.printRange_Click);
            // 
            // billNoFrom
            // 
            this.billNoFrom.Location = new System.Drawing.Point(284, 212);
            this.billNoFrom.Name = "billNoFrom";
            this.billNoFrom.Size = new System.Drawing.Size(110, 26);
            this.billNoFrom.TabIndex = 8;
            this.billNoFrom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.billNoFrom_KeyPress);
            // 
            // billNoTo
            // 
            this.billNoTo.Location = new System.Drawing.Point(425, 212);
            this.billNoTo.Name = "billNoTo";
            this.billNoTo.Size = new System.Drawing.Size(123, 26);
            this.billNoTo.TabIndex = 9;
            this.billNoTo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.billNoTo_KeyPress);
            // 
            // printAll
            // 
            this.printAll.Location = new System.Drawing.Point(284, 313);
            this.printAll.Name = "printAll";
            this.printAll.Size = new System.Drawing.Size(95, 39);
            this.printAll.TabIndex = 10;
            this.printAll.Text = "Print All";
            this.printAll.UseVisualStyleBackColor = true;
            this.printAll.Click += new System.EventHandler(this.printAll_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(400, 212);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(19, 20);
            this.label4.TabIndex = 11;
            this.label4.Text = "--";
            // 
            // GPTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.printAll);
            this.Controls.Add(this.billNoTo);
            this.Controls.Add(this.billNoFrom);
            this.Controls.Add(this.printRange);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ErrorMsg);
            this.Controls.Add(this.billNum);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.excelPath);
            this.Name = "GPTool";
            this.Text = "GP Tool";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox excelPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox billNum;
        private System.Windows.Forms.Label ErrorMsg;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.PrintPreviewDialog printPreviewDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button printRange;
        private System.Windows.Forms.TextBox billNoFrom;
        private System.Windows.Forms.TextBox billNoTo;
        private System.Windows.Forms.Button printAll;
        private System.Windows.Forms.Label label4;
    }
}

