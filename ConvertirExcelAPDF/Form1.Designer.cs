namespace ConvertirExcelAPDF
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnSelectFiles = new Button();
            btnConvertToPdf = new Button();
            listBoxFiles = new ListBox();
            pictureBox = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)pictureBox).BeginInit();
            SuspendLayout();
            // 
            // btnSelectFiles
            // 
            btnSelectFiles.BackColor = Color.BlueViolet;
            btnSelectFiles.FlatStyle = FlatStyle.Flat;
            btnSelectFiles.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            btnSelectFiles.ForeColor = Color.White;
            btnSelectFiles.Location = new Point(571, 12);
            btnSelectFiles.Name = "btnSelectFiles";
            btnSelectFiles.Size = new Size(125, 23);
            btnSelectFiles.TabIndex = 0;
            btnSelectFiles.Text = "Seleccionar";
            btnSelectFiles.UseVisualStyleBackColor = false;
            btnSelectFiles.Click += btnSelectFiles_Click;
            // 
            // btnConvertToPdf
            // 
            btnConvertToPdf.BackColor = Color.BlueViolet;
            btnConvertToPdf.FlatStyle = FlatStyle.Flat;
            btnConvertToPdf.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            btnConvertToPdf.ForeColor = Color.White;
            btnConvertToPdf.Location = new Point(571, 54);
            btnConvertToPdf.Name = "btnConvertToPdf";
            btnConvertToPdf.Size = new Size(125, 23);
            btnConvertToPdf.TabIndex = 1;
            btnConvertToPdf.Text = "Convertir";
            btnConvertToPdf.UseVisualStyleBackColor = false;
            btnConvertToPdf.Click += btnConvertToPdf_Click;
            // 
            // listBoxFiles
            // 
            listBoxFiles.BackColor = Color.BlueViolet;
            listBoxFiles.ForeColor = Color.White;
            listBoxFiles.FormattingEnabled = true;
            listBoxFiles.ItemHeight = 15;
            listBoxFiles.Location = new Point(12, 12);
            listBoxFiles.Name = "listBoxFiles";
            listBoxFiles.Size = new Size(536, 169);
            listBoxFiles.TabIndex = 2;
            // 
            // pictureBox
            // 
            pictureBox.Image = (Image)resources.GetObject("pictureBox.Image");
            pictureBox.Location = new Point(571, 92);
            pictureBox.Name = "pictureBox";
            pictureBox.Size = new Size(125, 89);
            pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox.TabIndex = 3;
            pictureBox.TabStop = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Purple;
            ClientSize = new Size(719, 205);
            Controls.Add(pictureBox);
            Controls.Add(listBoxFiles);
            Controls.Add(btnConvertToPdf);
            Controls.Add(btnSelectFiles);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)pictureBox).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private Button btnSelectFiles;
        private Button btnConvertToPdf;
        private ListBox listBoxFiles;
        private PictureBox pictureBox;
    }
}