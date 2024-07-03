namespace GeneradorReportes
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private Button buttonSelectExcel;
        private Button buttonProcess;
        private TextBox textBoxExcel;
        private SaveFileDialog saveFileDialog; // Añadir esta línea

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            buttonSelectExcel = new Button();
            buttonProcess = new Button();
            textBoxExcel = new TextBox();
            saveFileDialog = new SaveFileDialog();
            textBoxPeriodo = new TextBox();
            textBoxNombre = new TextBox();
            textBoxTitulo = new TextBox();
            textBoxPeriodoInforme = new TextBox();
            textBoxRutaExcel = new TextBox();
            pictureBoxRuedaCargando = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)pictureBoxRuedaCargando).BeginInit();
            SuspendLayout();
            // 
            // buttonSelectExcel
            // 
            buttonSelectExcel.BackColor = Color.BlueViolet;
            buttonSelectExcel.ForeColor = Color.White;
            buttonSelectExcel.Location = new Point(467, 26);
            buttonSelectExcel.Margin = new Padding(4, 3, 4, 3);
            buttonSelectExcel.Name = "buttonSelectExcel";
            buttonSelectExcel.Size = new Size(126, 27);
            buttonSelectExcel.TabIndex = 0;
            buttonSelectExcel.Text = "Seleccionar Excel";
            buttonSelectExcel.UseVisualStyleBackColor = false;
            buttonSelectExcel.Click += button1_Click;
            // 
            // buttonProcess
            // 
            buttonProcess.BackColor = Color.BlueViolet;
            buttonProcess.ForeColor = Color.White;
            buttonProcess.Location = new Point(467, 92);
            buttonProcess.Margin = new Padding(4, 3, 4, 3);
            buttonProcess.Name = "buttonProcess";
            buttonProcess.Size = new Size(126, 27);
            buttonProcess.TabIndex = 1;
            buttonProcess.Text = "Procesar";
            buttonProcess.UseVisualStyleBackColor = false;
            buttonProcess.Click += buttonProcesar_Click;
            // 
            // textBoxExcel
            // 
            textBoxExcel.BackColor = Color.BlueViolet;
            textBoxExcel.ForeColor = Color.White;
            textBoxExcel.Location = new Point(29, 29);
            textBoxExcel.Margin = new Padding(4, 3, 4, 3);
            textBoxExcel.Name = "textBoxExcel";
            textBoxExcel.Size = new Size(408, 23);
            textBoxExcel.TabIndex = 2;
            // 
            // textBoxPeriodo
            // 
            textBoxPeriodo.BackColor = Color.BlueViolet;
            textBoxPeriodo.ForeColor = Color.White;
            textBoxPeriodo.Location = new Point(29, 95);
            textBoxPeriodo.Name = "textBoxPeriodo";
            textBoxPeriodo.Size = new Size(408, 23);
            textBoxPeriodo.TabIndex = 3;
            // 
            // textBoxNombre
            // 
            textBoxNombre.BackColor = Color.BlueViolet;
            textBoxNombre.ForeColor = Color.White;
            textBoxNombre.Location = new Point(29, 168);
            textBoxNombre.Name = "textBoxNombre";
            textBoxNombre.Size = new Size(408, 23);
            textBoxNombre.TabIndex = 4;
            // 
            // textBoxTitulo
            // 
            textBoxTitulo.BackColor = Color.Purple;
            textBoxTitulo.BorderStyle = BorderStyle.None;
            textBoxTitulo.Font = new Font("Segoe UI Black", 9.75F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxTitulo.ForeColor = Color.White;
            textBoxTitulo.Location = new Point(29, 146);
            textBoxTitulo.Name = "textBoxTitulo";
            textBoxTitulo.Size = new Size(120, 18);
            textBoxTitulo.TabIndex = 5;
            textBoxTitulo.Text = "Título del informe";
            // 
            // textBoxPeriodoInforme
            // 
            textBoxPeriodoInforme.BackColor = Color.Purple;
            textBoxPeriodoInforme.BorderStyle = BorderStyle.None;
            textBoxPeriodoInforme.Font = new Font("Segoe UI Black", 9.75F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxPeriodoInforme.ForeColor = Color.White;
            textBoxPeriodoInforme.Location = new Point(29, 73);
            textBoxPeriodoInforme.Name = "textBoxPeriodoInforme";
            textBoxPeriodoInforme.Size = new Size(100, 18);
            textBoxPeriodoInforme.TabIndex = 6;
            textBoxPeriodoInforme.Text = "Período";
            // 
            // textBoxRutaExcel
            // 
            textBoxRutaExcel.BackColor = Color.Purple;
            textBoxRutaExcel.BorderStyle = BorderStyle.None;
            textBoxRutaExcel.Font = new Font("Segoe UI Black", 9.75F, FontStyle.Bold, GraphicsUnit.Point);
            textBoxRutaExcel.ForeColor = Color.White;
            textBoxRutaExcel.Location = new Point(29, 7);
            textBoxRutaExcel.Name = "textBoxRutaExcel";
            textBoxRutaExcel.Size = new Size(120, 18);
            textBoxRutaExcel.TabIndex = 7;
            textBoxRutaExcel.Text = "Ruta del archivo";
            // 
            // pictureBoxRuedaCargando
            // 
            pictureBoxRuedaCargando.Image = (Image)resources.GetObject("pictureBoxRuedaCargando.Image");
            pictureBoxRuedaCargando.Location = new Point(483, 135);
            pictureBoxRuedaCargando.Name = "pictureBoxRuedaCargando";
            pictureBoxRuedaCargando.Size = new Size(91, 66);
            pictureBoxRuedaCargando.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBoxRuedaCargando.TabIndex = 8;
            pictureBoxRuedaCargando.TabStop = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Purple;
            ClientSize = new Size(618, 213);
            Controls.Add(pictureBoxRuedaCargando);
            Controls.Add(textBoxRutaExcel);
            Controls.Add(textBoxPeriodoInforme);
            Controls.Add(textBoxTitulo);
            Controls.Add(textBoxNombre);
            Controls.Add(textBoxPeriodo);
            Controls.Add(textBoxExcel);
            Controls.Add(buttonProcess);
            Controls.Add(buttonSelectExcel);
            Margin = new Padding(4, 3, 4, 3);
            Name = "Form1";
            Text = "Report Generator";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)pictureBoxRuedaCargando).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        private TextBox textBoxPeriodo;
        private TextBox textBoxNombre;
        private TextBox textBoxTitulo;
        private TextBox textBoxPeriodoInforme;
        private TextBox textBoxRutaExcel;
        private PictureBox pictureBoxRuedaCargando;
    }
}
