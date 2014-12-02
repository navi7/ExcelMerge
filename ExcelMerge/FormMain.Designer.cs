namespace ExcelMerge
{
    partial class FormMain
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
            this.groupBoxInput = new System.Windows.Forms.GroupBox();
            this.buttonOpenFile2 = new System.Windows.Forms.Button();
            this.buttonOpenFile1 = new System.Windows.Forms.Button();
            this.textBoxInput2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxInput1 = new System.Windows.Forms.TextBox();
            this.groupBoxResult = new System.Windows.Forms.GroupBox();
            this.dataGridViewResult = new System.Windows.Forms.DataGridView();
            this.buttonSave = new System.Windows.Forms.Button();
            this.buttonProcess = new System.Windows.Forms.Button();
            this.groupBoxInput.SuspendLayout();
            this.groupBoxResult.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResult)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBoxInput
            // 
            this.groupBoxInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxInput.Controls.Add(this.buttonOpenFile2);
            this.groupBoxInput.Controls.Add(this.buttonOpenFile1);
            this.groupBoxInput.Controls.Add(this.textBoxInput2);
            this.groupBoxInput.Controls.Add(this.label2);
            this.groupBoxInput.Controls.Add(this.label1);
            this.groupBoxInput.Controls.Add(this.textBoxInput1);
            this.groupBoxInput.Location = new System.Drawing.Point(13, 14);
            this.groupBoxInput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBoxInput.Name = "groupBoxInput";
            this.groupBoxInput.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBoxInput.Size = new System.Drawing.Size(664, 108);
            this.groupBoxInput.TabIndex = 0;
            this.groupBoxInput.TabStop = false;
            this.groupBoxInput.Text = "Ulazne datoteke";
            // 
            // buttonOpenFile2
            // 
            this.buttonOpenFile2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOpenFile2.Location = new System.Drawing.Point(624, 65);
            this.buttonOpenFile2.Name = "buttonOpenFile2";
            this.buttonOpenFile2.Size = new System.Drawing.Size(32, 29);
            this.buttonOpenFile2.TabIndex = 5;
            this.buttonOpenFile2.Text = "...";
            this.buttonOpenFile2.UseVisualStyleBackColor = true;
            this.buttonOpenFile2.Click += new System.EventHandler(this.buttonOpenFile2_Click);
            // 
            // buttonOpenFile1
            // 
            this.buttonOpenFile1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOpenFile1.Location = new System.Drawing.Point(625, 26);
            this.buttonOpenFile1.Name = "buttonOpenFile1";
            this.buttonOpenFile1.Size = new System.Drawing.Size(32, 29);
            this.buttonOpenFile1.TabIndex = 4;
            this.buttonOpenFile1.Text = "...";
            this.buttonOpenFile1.UseVisualStyleBackColor = true;
            this.buttonOpenFile1.Click += new System.EventHandler(this.buttonOpenFile1_Click);
            // 
            // textBoxInput2
            // 
            this.textBoxInput2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxInput2.Location = new System.Drawing.Point(252, 65);
            this.textBoxInput2.Name = "textBoxInput2";
            this.textBoxInput2.Size = new System.Drawing.Size(366, 29);
            this.textBoxInput2.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(239, 21);
            this.label2.TabIndex = 2;
            this.label2.Text = "Datoteka sa stanjem na skladištu:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(166, 21);
            this.label1.TabIndex = 1;
            this.label1.Text = "Datoteka s potrebama:";
            // 
            // textBoxInput1
            // 
            this.textBoxInput1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxInput1.Location = new System.Drawing.Point(182, 26);
            this.textBoxInput1.Name = "textBoxInput1";
            this.textBoxInput1.Size = new System.Drawing.Size(436, 29);
            this.textBoxInput1.TabIndex = 0;
            // 
            // groupBoxResult
            // 
            this.groupBoxResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxResult.Controls.Add(this.dataGridViewResult);
            this.groupBoxResult.Location = new System.Drawing.Point(13, 128);
            this.groupBoxResult.Name = "groupBoxResult";
            this.groupBoxResult.Size = new System.Drawing.Size(826, 412);
            this.groupBoxResult.TabIndex = 1;
            this.groupBoxResult.TabStop = false;
            this.groupBoxResult.Text = "Rezultat";
            // 
            // dataGridViewResult
            // 
            this.dataGridViewResult.AllowUserToAddRows = false;
            this.dataGridViewResult.AllowUserToDeleteRows = false;
            this.dataGridViewResult.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridViewResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewResult.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewResult.Location = new System.Drawing.Point(3, 25);
            this.dataGridViewResult.Name = "dataGridViewResult";
            this.dataGridViewResult.Size = new System.Drawing.Size(820, 384);
            this.dataGridViewResult.TabIndex = 0;
            this.dataGridViewResult.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dataGridViewResult_CellFormatting);
            // 
            // buttonSave
            // 
            this.buttonSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSave.Enabled = false;
            this.buttonSave.Image = global::ExcelMerge.Properties.Resources.save;
            this.buttonSave.Location = new System.Drawing.Point(684, 546);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(155, 36);
            this.buttonSave.TabIndex = 6;
            this.buttonSave.Text = "Spremi...";
            this.buttonSave.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // buttonProcess
            // 
            this.buttonProcess.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonProcess.Enabled = false;
            this.buttonProcess.Image = global::ExcelMerge.Properties.Resources.wizard;
            this.buttonProcess.Location = new System.Drawing.Point(684, 28);
            this.buttonProcess.Name = "buttonProcess";
            this.buttonProcess.Size = new System.Drawing.Size(155, 94);
            this.buttonProcess.TabIndex = 5;
            this.buttonProcess.Text = "Učitaj datoteke!";
            this.buttonProcess.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.buttonProcess.UseVisualStyleBackColor = true;
            this.buttonProcess.Click += new System.EventHandler(this.buttonProcess_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(851, 585);
            this.Controls.Add(this.buttonSave);
            this.Controls.Add(this.buttonProcess);
            this.Controls.Add(this.groupBoxResult);
            this.Controls.Add(this.groupBoxInput);
            this.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormMain";
            this.Text = "Excel merge";
            this.groupBoxInput.ResumeLayout(false);
            this.groupBoxInput.PerformLayout();
            this.groupBoxResult.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewResult)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxInput;
        private System.Windows.Forms.GroupBox groupBoxResult;
        private System.Windows.Forms.DataGridView dataGridViewResult;
        private System.Windows.Forms.TextBox textBoxInput2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxInput1;
        private System.Windows.Forms.Button buttonOpenFile1;
        private System.Windows.Forms.Button buttonOpenFile2;
        private System.Windows.Forms.Button buttonProcess;
        private System.Windows.Forms.Button buttonSave;
    }
}

