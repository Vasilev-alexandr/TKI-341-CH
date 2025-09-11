using System.Drawing;
using System.Windows.Forms;

namespace VasilevPracticeExcel
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
            PreviewDGV = new DataGridView();
            Column1 = new DataGridViewTextBoxColumn();
            Column2 = new DataGridViewTextBoxColumn();
            Column3 = new DataGridViewTextBoxColumn();
            Column4 = new DataGridViewTextBoxColumn();
            Column5 = new DataGridViewTextBoxColumn();
            Column6 = new DataGridViewTextBoxColumn();
            Column7 = new DataGridViewTextBoxColumn();
            Column8 = new DataGridViewTextBoxColumn();
            Column9 = new DataGridViewTextBoxColumn();
            buttonSave = new Button();
            label1 = new Label();
            ((System.ComponentModel.ISupportInitialize)PreviewDGV).BeginInit();
            SuspendLayout();
            // 
            // PreviewDGV
            // 
            PreviewDGV.AllowUserToAddRows = false;
            PreviewDGV.AllowUserToDeleteRows = false;
            PreviewDGV.AllowUserToOrderColumns = true;
            PreviewDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            PreviewDGV.Columns.AddRange(new DataGridViewColumn[] { Column1, Column2, Column3, Column4, Column5, Column6, Column7, Column8, Column9 });
            PreviewDGV.Location = new Point(12, 27);
            PreviewDGV.Name = "PreviewDGV";
            PreviewDGV.ReadOnly = true;
            PreviewDGV.Size = new Size(943, 260);
            PreviewDGV.TabIndex = 1;
            // 
            // Column1
            // 
            Column1.HeaderText = "Дата записи";
            Column1.Name = "Column1";
            Column1.ReadOnly = true;
            // 
            // Column2
            // 
            Column2.HeaderText = "Номер документа";
            Column2.Name = "Column2";
            Column2.ReadOnly = true;
            // 
            // Column3
            // 
            Column3.HeaderText = "Номер по порядку";
            Column3.Name = "Column3";
            Column3.ReadOnly = true;
            // 
            // Column4
            // 
            Column4.HeaderText = "Получено/отправлено";
            Column4.Name = "Column4";
            Column4.ReadOnly = true;
            // 
            // Column5
            // 
            Column5.HeaderText = "Единица продукции";
            Column5.Name = "Column5";
            Column5.ReadOnly = true;
            // 
            // Column6
            // 
            Column6.HeaderText = "Приход";
            Column6.Name = "Column6";
            Column6.ReadOnly = true;
            // 
            // Column7
            // 
            Column7.HeaderText = "Расход";
            Column7.Name = "Column7";
            Column7.ReadOnly = true;
            // 
            // Column8
            // 
            Column8.HeaderText = "Остаток";
            Column8.Name = "Column8";
            Column8.ReadOnly = true;
            // 
            // Column9
            // 
            Column9.HeaderText = "Дата, подпись";
            Column9.Name = "Column9";
            Column9.ReadOnly = true;
            // 
            // buttonSave
            // 
            buttonSave.Location = new Point(843, 293);
            buttonSave.Name = "buttonSave";
            buttonSave.Size = new Size(111, 23);
            buttonSave.TabIndex = 0;
            buttonSave.Text = "Выгрузка в Excel";
            buttonSave.UseVisualStyleBackColor = true;
            buttonSave.Click += buttonSave_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 9);
            label1.Name = "label1";
            label1.Size = new Size(90, 15);
            label1.TabIndex = 2;
            label1.Text = "Предпросмотр";
            // 
            // ExcelAutomationForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(966, 324);
            Controls.Add(label1);
            Controls.Add(PreviewDGV);
            Controls.Add(buttonSave);
            Name = "ExcelAutomationForm";
            Text = "Задание №2 выполнил: Васильев А.С., Номер варианта: 3 Дата выполнения: 15/08/2025";
            ((System.ComponentModel.ISupportInitialize)PreviewDGV).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button buttonSave;
        private DataGridView PreviewDGV;
        private Label label1;
        private DataGridViewTextBoxColumn Column1;
        private DataGridViewTextBoxColumn Column2;
        private DataGridViewTextBoxColumn Column3;
        private DataGridViewTextBoxColumn Column4;
        private DataGridViewTextBoxColumn Column5;
        private DataGridViewTextBoxColumn Column6;
        private DataGridViewTextBoxColumn Column7;
        private DataGridViewTextBoxColumn Column8;
        private DataGridViewTextBoxColumn Column9;
    }
}
