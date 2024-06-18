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
            buttonSelectExcel = new Button();
            buttonProcess = new Button();
            textBoxExcel = new TextBox();
            saveFileDialog = new SaveFileDialog();
            textBoxPeriodo = new TextBox();
            textBoxNombre = new TextBox();
            SuspendLayout();
            // 
            // buttonSelectExcel
            // 
            buttonSelectExcel.Location = new Point(467, 29);
            buttonSelectExcel.Margin = new Padding(4, 3, 4, 3);
            buttonSelectExcel.Name = "buttonSelectExcel";
            buttonSelectExcel.Size = new Size(88, 27);
            buttonSelectExcel.TabIndex = 0;
            buttonSelectExcel.Text = "Select Excel";
            buttonSelectExcel.UseVisualStyleBackColor = true;
            buttonSelectExcel.Click += button1_Click;
            // 
            // buttonProcess
            // 
            buttonProcess.Location = new Point(467, 73);
            buttonProcess.Margin = new Padding(4, 3, 4, 3);
            buttonProcess.Name = "buttonProcess";
            buttonProcess.Size = new Size(88, 27);
            buttonProcess.TabIndex = 1;
            buttonProcess.Text = "Process";
            buttonProcess.UseVisualStyleBackColor = true;
            buttonProcess.Click += buttonProcesar_Click;
            // 
            // textBoxExcel
            // 
            textBoxExcel.Location = new Point(29, 29);
            textBoxExcel.Margin = new Padding(4, 3, 4, 3);
            textBoxExcel.Name = "textBoxExcel";
            textBoxExcel.Size = new Size(408, 23);
            textBoxExcel.TabIndex = 2;
            // 
            // textBoxPeriodo
            // 
            textBoxPeriodo.Location = new Point(29, 73);
            textBoxPeriodo.Name = "textBoxPeriodo";
            textBoxPeriodo.Size = new Size(408, 23);
            textBoxPeriodo.TabIndex = 3;
            textBoxPeriodo.Text = "Definir período en formato número, ejemplo junio es \"06\"";
            // 
            // textBoxNombre
            // 
            textBoxNombre.Location = new Point(29, 112);
            textBoxNombre.Name = "textBoxNombre";
            textBoxNombre.Size = new Size(408, 23);
            textBoxNombre.TabIndex = 4;
            textBoxNombre.Text = "Titulo del informe, ejemplo \"ESCOBAR\"";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(606, 147);
            Controls.Add(textBoxNombre);
            Controls.Add(textBoxPeriodo);
            Controls.Add(textBoxExcel);
            Controls.Add(buttonProcess);
            Controls.Add(buttonSelectExcel);
            Margin = new Padding(4, 3, 4, 3);
            Name = "Form1";
            Text = "Report Generator";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        private TextBox textBoxPeriodo;
        private TextBox textBoxNombre;
    }
}
