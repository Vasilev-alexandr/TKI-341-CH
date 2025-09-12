using System.Drawing;
using System.Windows.Forms;

namespace WordAutomation
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

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
            txtMinistry = new TextBox();
            txtUniversityType = new TextBox();
            txtUniversity = new TextBox();
            txtInstitute = new TextBox();
            txtDepartment = new TextBox();
            cmbDocumentType = new ComboBox();
            txtLink1 = new TextBox();
            cmbWorkType = new ComboBox();
            txtLink2 = new TextBox();
            txtDiscipline = new TextBox();
            txtLink3 = new TextBox();
            txtGroup = new TextBox();
            txtAuthor = new TextBox();
            txtChecker = new TextBox();
            txtCityYear = new TextBox();
            btnSave = new Button();
            btnCreate = new Button();
            cmbWorkNumber = new ComboBox();
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            label5 = new Label();
            label6 = new Label();
            label7 = new Label();
            label8 = new Label();
            label9 = new Label();
            label10 = new Label();
            label11 = new Label();
            label12 = new Label();
            label13 = new Label();
            label14 = new Label();
            label15 = new Label();
            label16 = new Label();
            trackBarLeftMargin = new TrackBar();
            label17 = new Label();
            lblMarginValue = new Label();
            ((System.ComponentModel.ISupportInitialize)trackBarLeftMargin).BeginInit();
            SuspendLayout();
            
            txtMinistry.Location = new Point(200, 51);
            txtMinistry.Name = "txtMinistry";
            txtMinistry.Size = new Size(300, 23);
            txtMinistry.TabIndex = 0;
            txtMinistry.Text = "Министерство транспорта Российской Федерации";

            txtUniversityType.Location = new Point(200, 80);
            txtUniversityType.Name = "txtUniversityType";
            txtUniversityType.Size = new Size(300, 23);
            txtUniversityType.TabIndex = 1;
            txtUniversityType.Text = "Федеральное государственное бюджетное образовательное учреждение высшего образования";

            txtUniversity.Location = new Point(200, 110);
            txtUniversity.Name = "txtUniversity";
            txtUniversity.Size = new Size(300, 23);
            txtUniversity.TabIndex = 2;
            txtUniversity.Text = "«Российский университет транспорта» (РУТ (МИИТ))";

            txtInstitute.Location = new Point(200, 140);
            txtInstitute.Name = "txtInstitute";
            txtInstitute.Size = new Size(300, 23);
            txtInstitute.TabIndex = 3;
            txtInstitute.Text = "Институт транспортной техники и систем управления";

            txtDepartment.Location = new Point(200, 170);
            txtDepartment.Name = "txtDepartment";
            txtDepartment.Size = new Size(300, 23);
            txtDepartment.TabIndex = 4;
            txtDepartment.Text = "Кафедра «Управление и защита информации»";
 
            cmbDocumentType.FormattingEnabled = true;
            cmbDocumentType.Location = new Point(200, 200);
            cmbDocumentType.Name = "cmbDocumentType";
            cmbDocumentType.Size = new Size(300, 23);
            cmbDocumentType.TabIndex = 5;

            txtLink1.Location = new Point(200, 230);
            txtLink1.Name = "txtLink1";
            txtLink1.Size = new Size(300, 23);
            txtLink1.TabIndex = 6;
            txtLink1.Text = "по практике";

            cmbWorkType.FormattingEnabled = true;
            cmbWorkType.Location = new Point(200, 260);
            cmbWorkType.Name = "cmbWorkType";
            cmbWorkType.Size = new Size(300, 23);
            cmbWorkType.TabIndex = 7;
 
            txtLink2.Location = new Point(200, 290);
            txtLink2.Name = "txtLink2";
            txtLink2.Size = new Size(300, 23);
            txtLink2.TabIndex = 8;
            txtLink2.Text = "по дисциплине";

            txtDiscipline.Location = new Point(200, 320);
            txtDiscipline.Name = "txtDiscipline";
            txtDiscipline.Size = new Size(300, 23);
            txtDiscipline.TabIndex = 9;
            txtDiscipline.Text = "«Алгоритмизация и технологии программирования»";

            txtLink3.Location = new Point(200, 350);
            txtLink3.Name = "txtLink3";
            txtLink3.Size = new Size(300, 23);
            txtLink3.TabIndex = 10;
            txtLink3.Text = "на тему";
  
            txtGroup.Location = new Point(200, 380);
            txtGroup.Name = "txtGroup";
            txtGroup.Size = new Size(300, 23);
            txtGroup.TabIndex = 11;
            txtGroup.Text = "Выполнил: ст. гр. ТКИ-341";

            txtAuthor.Location = new Point(200, 410);
            txtAuthor.Name = "txtAuthor";
            txtAuthor.Size = new Size(300, 23);
            txtAuthor.TabIndex = 12;
            txtAuthor.Text = "Васильев Александр Сергеевич";
 
            txtChecker.Location = new Point(200, 440);
            txtChecker.Name = "txtChecker";
            txtChecker.Size = new Size(300, 23);
            txtChecker.TabIndex = 13;
            txtChecker.Text = "Проверил: доц. к.т.н. Сафроноа А.И.";

            txtCityYear.Location = new Point(200, 470);
            txtCityYear.Name = "txtCityYear";
            txtCityYear.Size = new Size(300, 23);
            txtCityYear.TabIndex = 14;
            txtCityYear.Text = "Москва - 2025";

            btnSave.Location = new Point(200, 550);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(140, 30);
            btnSave.TabIndex = 15;
            btnSave.Text = "Сохранить";
            btnSave.UseVisualStyleBackColor = true;
            btnSave.Click += btnSave_Click;

            btnCreate.Location = new Point(360, 550);
            btnCreate.Name = "btnCreate";
            btnCreate.Size = new Size(140, 30);
            btnCreate.TabIndex = 16;
            btnCreate.Text = "Создать";
            btnCreate.UseVisualStyleBackColor = true;
            btnCreate.Click += btnCreate_Click;

            cmbWorkNumber.FormattingEnabled = true;
            cmbWorkNumber.Location = new Point(200, 260);
            cmbWorkNumber.Name = "cmbWorkNumber";
            cmbWorkNumber.Size = new Size(300, 23);
            cmbWorkNumber.TabIndex = 17;
            cmbWorkNumber.Visible = false;

            label1.AutoSize = true;
            label1.Location = new Point(20, 53);
            label1.Name = "label1";
            label1.Size = new Size(154, 15);
            label1.TabIndex = 18;
            label1.Text = "Министерство/Ведомство:";

            label2.AutoSize = true;
            label2.Location = new Point(20, 83);
            label2.Name = "label2";
            label2.Size = new Size(100, 15);
            label2.TabIndex = 19;
            label2.Text = "Тип учреждения:";

            label3.AutoSize = true;
            label3.Location = new Point(20, 113);
            label3.Name = "label3";
            label3.Size = new Size(105, 15);
            label3.TabIndex = 20;

            label4.AutoSize = true;
            label4.Location = new Point(20, 143);
            label4.Name = "label4";
            label4.Size = new Size(60, 15);
            label4.TabIndex = 21;
            label4.Text = "Институт:";
 
            label5.AutoSize = true;
            label5.Location = new Point(20, 173);
            label5.Name = "label5";
            label5.Size = new Size(57, 15);
            label5.TabIndex = 22;
            label5.Text = "Кафедра:";

            label6.AutoSize = true;
            label6.Location = new Point(20, 203);
            label6.Name = "label6";
            label6.Size = new Size(91, 15);
            label6.TabIndex = 23;
            label6.Text = "Тип документа:";

            label7.AutoSize = true;
            label7.Location = new Point(20, 233);
            label7.Name = "label7";
            label7.Size = new Size(90, 15);
            label7.TabIndex = 24;
            label7.Text = "Связка 1 (по...):";

            label8.AutoSize = true;
            label8.Location = new Point(20, 263);
            label8.Name = "label8";
            label8.Size = new Size(74, 15);
            label8.TabIndex = 25;
            label8.Text = "Вид работы:";

            label9.AutoSize = true;
            label9.Location = new Point(20, 293);
            label9.Name = "label9";
            label9.Size = new Size(90, 15);
            label9.TabIndex = 26;
            label9.Text = "Связка 2 (по...):";

            label10.AutoSize = true;
            label10.Location = new Point(20, 323);
            label10.Name = "label10";
            label10.Size = new Size(79, 15);
            label10.TabIndex = 27;
            label10.Text = "Дисциплина:";
  
            label11.AutoSize = true;
            label11.Location = new Point(20, 353);
            label11.Name = "label11";
            label11.Size = new Size(118, 15);
            label11.TabIndex = 28;
            label11.Text = "Связка 3 (на тему...):";
    
            label12.AutoSize = true;
            label12.Location = new Point(20, 383);
            label12.Name = "label12";
            label12.Size = new Size(49, 15);
            label12.TabIndex = 29;
            label12.Text = "Группа:";
          
            label13.AutoSize = true;
            label13.Location = new Point(20, 413);
            label13.Name = "label13";
            label13.Size = new Size(43, 15);
            label13.TabIndex = 30;
            label13.Text = "Автор:";
       
            label14.AutoSize = true;
            label14.Location = new Point(20, 443);
            label14.Name = "label14";
            label14.Size = new Size(66, 15);
            label14.TabIndex = 31;
            label14.Text = "Проверил:";
          
            label15.AutoSize = true;
            label15.Location = new Point(20, 473);
            label15.Name = "label15";
            label15.Size = new Size(100, 15);
            label15.TabIndex = 32;
            label15.Text = "Город и год (г-г):";
           
            label16.AutoSize = true;
            label16.Location = new Point(20, 263);
            label16.Name = "label16";
            label16.Size = new Size(0, 15);
            label16.TabIndex = 33;
         
            trackBarLeftMargin.Location = new Point(200, 510);
            trackBarLeftMargin.Maximum = 200;
            trackBarLeftMargin.Name = "trackBarLeftMargin";
            trackBarLeftMargin.Size = new Size(250, 45);
            trackBarLeftMargin.TabIndex = 34;
            trackBarLeftMargin.Scroll += trackBarLeftMargin_Scroll;
         
            label17.AutoSize = true;
            label17.Location = new Point(20, 513);
            label17.Name = "label17";
            label17.Size = new Size(174, 15);
            label17.TabIndex = 35;
            label17.Text = "Смещение левой границы :";
           
            lblMarginValue.AutoSize = true;
            lblMarginValue.Location = new Point(460, 513);
            lblMarginValue.Name = "lblMarginValue";
            lblMarginValue.Size = new Size(13, 15);
            lblMarginValue.TabIndex = 36;
            lblMarginValue.Text = "0";
           
            ClientSize = new Size(534, 600);
            Controls.Add(lblMarginValue);
            Controls.Add(label17);
            Controls.Add(trackBarLeftMargin);
            Controls.Add(label16);
            Controls.Add(label15);
            Controls.Add(label14);
            Controls.Add(label13);
            Controls.Add(label12);
            Controls.Add(label11);
            Controls.Add(label10);
            Controls.Add(label9);
            Controls.Add(label8);
            Controls.Add(label7);
            Controls.Add(label6);
            Controls.Add(label5);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(cmbWorkNumber);
            Controls.Add(btnCreate);
            Controls.Add(btnSave);
            Controls.Add(txtCityYear);
            Controls.Add(txtChecker);
            Controls.Add(txtAuthor);
            Controls.Add(txtGroup);
            Controls.Add(txtLink3);
            Controls.Add(txtDiscipline);
            Controls.Add(txtLink2);
            Controls.Add(cmbWorkType);
            Controls.Add(txtLink1);
            Controls.Add(cmbDocumentType);
            Controls.Add(txtDepartment);
            Controls.Add(txtInstitute);
            Controls.Add(txtUniversity);
            Controls.Add(txtUniversityType);
            Controls.Add(txtMinistry);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            ((System.ComponentModel.ISupportInitialize)trackBarLeftMargin).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        private TextBox txtMinistry;
        private TextBox txtUniversityType;
        private TextBox txtUniversity;
        private TextBox txtInstitute;
        private TextBox txtDepartment;
        private ComboBox cmbDocumentType;
        private TextBox txtLink1;
        private ComboBox cmbWorkType;
        private TextBox txtLink2;
        private TextBox txtDiscipline;
        private TextBox txtLink3;
        private TextBox txtGroup;
        private TextBox txtAuthor;
        private TextBox txtChecker;
        private TextBox txtCityYear;
        private Button btnSave;
        private Button btnCreate;
        private ComboBox cmbWorkNumber;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private Label label9;
        private Label label10;
        private Label label11;
        private Label label12;
        private Label label13;
        private Label label14;
        private Label label15;
        private Label label16;
        private TrackBar trackBarLeftMargin;
        private Label label17;
        private Label lblMarginValue;
    }
}
