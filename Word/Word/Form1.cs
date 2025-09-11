using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Xml;

namespace WordAutomation
{
    public partial class Form1 : Form
    {
        private string configFile = "config.txt";
        private string authorName = "Васильев А.С.";
        private int variantNumber = 3;

        public Form1()
        {
            InitializeComponent();
            SetupForm();
            LoadConfig();
        }

        private void SetupForm()
        {
            this.Text = $"Задание №5 выполнил: Васильев А.С.; Номер варианта: 3; Дата выполнения: {DateTime.Now:dd/MM/yyyy}";

            // Инициализация комбобоксов
            cmbDocumentType.Items.AddRange(new string[] { "отчёт", "реферат", "эссе", "курсовой проект", "курсовая работа", "доклад", "домашнее задание" });
            cmbWorkType.Items.AddRange(new string[] { "лабораторная работа", "практическая работа", "индивидуальное задание", "учебная практика", "производственная практика", "преддипломная практика" });
            cmbWorkNumber.Items.AddRange(Enumerable.Range(1, 10).Select(i => i.ToString()).ToArray());

            // Установка значений по умолчанию
            cmbDocumentType.SelectedIndex = 0;
            cmbWorkType.SelectedIndex = 0;
            cmbWorkNumber.SelectedIndex = 0;
        }

        private void LoadConfig()
        {
            if (File.Exists(configFile))
            {
                try
                {
                    string[] lines = File.ReadAllLines(configFile);
                    if (lines.Length >= 15)
                    {
                        txtMinistry.Text = lines[0].Split(';')[0];
                        txtUniversityType.Text = lines[1].Split(';')[0];
                        txtUniversity.Text = lines[2].Split(';')[0];
                        txtInstitute.Text = lines[3].Split(';')[0];
                        txtDepartment.Text = lines[4].Split(';')[0];
                        cmbDocumentType.Text = lines[5].Split(';')[0];
                        txtLink1.Text = lines[6].Split(';')[0];
                        cmbWorkType.Text = lines[7].Split(';')[0];
                        txtLink2.Text = lines[8].Split(';')[0];
                        txtDiscipline.Text = lines[9].Split(';')[0];
                        txtLink3.Text = lines[10].Split(';')[0];
                        txtGroup.Text = lines[11].Split(';')[0];
                        txtAuthor.Text = lines[12].Split(';')[0];
                        txtChecker.Text = lines[13].Split(';')[0];
                        txtCityYear.Text = lines[14].Split(';')[0];
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка загрузки конфигурации: {ex.Message}");
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> lines = new List<string>
                {
                    $"{txtMinistry.Text};Times New Roman;14",
                    $"{txtUniversityType.Text};Times New Roman;14",
                    $"{txtUniversity.Text};Times New Roman;14",
                    $"{txtInstitute.Text};Times New Roman;14",
                    $"{txtDepartment.Text};Times New Roman;14",
                    $"{cmbDocumentType.Text};Times New Roman;14",
                    $"{txtLink1.Text};Times New Roman;12",
                    $"{cmbWorkType.Text};Times New Roman;12",
                    $"{txtLink2.Text};Times New Roman;12",
                    $"{txtDiscipline.Text};Times New Roman;12",
                    $"{txtLink3.Text};Times New Roman;12",
                    $"{txtGroup.Text};Times New Roman;12",
                    $"{txtAuthor.Text};Times New Roman;12",
                    $"{txtChecker.Text};Times New Roman;12",
                    $"{txtCityYear.Text};Times New Roman;12"
                };

                File.WriteAllLines(configFile, lines);
                MessageBox.Show("Конфигурация сохранена!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}");
            }
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            try
            {
                CreateWordDocument();
                MessageBox.Show("Документ создан успешно!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка создания документа: {ex.Message}");
            }
        }

        private void ConfigureDocumentStyles(Document doc, Microsoft.Office.Interop.Word.Application wordApp)
        {
            // Настройка стилей документа
            doc.Styles[WdBuiltinStyle.wdStyleNormal].Font.Name = "Times New Roman";
            doc.Styles[WdBuiltinStyle.wdStyleNormal].Font.Size = 12;

            // Настройка абзацев
            doc.Styles[WdBuiltinStyle.wdStyleNormal].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            doc.Styles[WdBuiltinStyle.wdStyleNormal].ParagraphFormat.FirstLineIndent = wordApp.CentimetersToPoints(1.25f);
            doc.Styles[WdBuiltinStyle.wdStyleNormal].ParagraphFormat.SpaceAfter = 6;
        }

        private void CreateWordDocument()
        {
            object missing = Missing.Value;
            object endOfDoc = "\\endofdoc";

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;

            try
            {
                Document doc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                doc.Activate();

                // Настройка страницы - стандартные поля для документа
                doc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(3f);
                doc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.5f);
                doc.PageSetup.TopMargin = wordApp.CentimetersToPoints(2f);
                doc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(2f);

                // Настройка стилей
                ConfigureDocumentStyles(doc, wordApp);

                // Добавление содержания
                AddTitleContent(doc, endOfDoc);

                // Сохранение
                string fileName = $"{DateTime.Now:yyyyMMdd}_Договор_коммерческой_концессии_{authorName}.docx";
                doc.SaveAs2(Path.Combine(System.Windows.Forms.Application.StartupPath, fileName));

                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка создания документа: {ex.Message}");
                wordApp.Quit();
            }
        }

        private void AddContractParties(Document doc, object endOfDoc)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;

            // Правообладатель
            Paragraph owner = doc.Content.Paragraphs.Add(ref range);
            owner.Range.Text = "«Правообладатель» ______, в лице ______, действующего на основании ______ и ______";
            owner.Range.Font.Name = "Times New Roman";
            owner.Range.Font.Size = 12;
            owner.Format.SpaceAfter = 6;
            owner.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            owner.Range.InsertParagraphAfter();

            // Пользователь
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph user = doc.Content.Paragraphs.Add(ref range);
            user.Range.Text = "«Пользователь» ______, в лице ______, действующего на основании ______";
            user.Range.Font.Name = "Times New Roman";
            user.Range.Font.Size = 12;
            user.Format.SpaceAfter = 6;
            user.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            user.Range.InsertParagraphAfter();

            // Подписание
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph sign = doc.Content.Paragraphs.Add(ref range);
            sign.Range.Text = "подписали этот договор коммерческой концессии на таких условиях.";
            sign.Range.Font.Name = "Times New Roman";
            sign.Range.Font.Size = 12;
            sign.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            sign.Format.SpaceAfter = 12;
            sign.Range.InsertParagraphAfter();
        }

        private void AddTitleContent(Document doc, object endOfDoc)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;

            // Заголовок ДОГОВОР
            Paragraph title = doc.Content.Paragraphs.Add(ref range);
            title.Range.Text = "ДОГОВОР";
            title.Range.Font.Name = "Times New Roman";
            title.Range.Font.Size = 16;
            title.Range.Font.Bold = 1;
            title.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            title.Format.SpaceAfter = 0;
            title.Range.InsertParagraphAfter();

            // Подзаголовок "коммерческой концессии"
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph subtitle = doc.Content.Paragraphs.Add(ref range);
            subtitle.Range.Text = "коммерческой концессии";
            subtitle.Range.Font.Name = "Times New Roman";
            subtitle.Range.Font.Size = 14;
            subtitle.Range.Font.Bold = 1;
            subtitle.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            subtitle.Format.SpaceAfter = 6; // Уменьшаем отступ
            subtitle.Range.InsertParagraphAfter();

            // Пустая строка
            AddEmptyParagraph(doc, endOfDoc);

            // Город и дата
            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph cityDate = doc.Content.Paragraphs.Add(ref range);
            cityDate.Range.Text = "г. ______";
            cityDate.Range.Font.Name = "Times New Roman";
            cityDate.Range.Font.Size = 12;
            cityDate.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            cityDate.Format.SpaceAfter = 0;
            cityDate.Range.InsertParagraphAfter();

            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph date = doc.Content.Paragraphs.Add(ref range);
            date.Range.Text = "«____» ______ 20__ г."; // Добавляем пробел перед "г."
            date.Range.Font.Name = "Times New Roman";
            date.Range.Font.Size = 12;
            date.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            date.Format.SpaceAfter = 12;
            date.Range.InsertParagraphAfter();

            // Стороны договора
            AddContractParties(doc, endOfDoc);

            // Текст договора
            AddContractText(doc, endOfDoc);
        }

        private void AddContractText(Document doc, object endOfDoc)
        {
            string[] clauses = {
        "1. Правообладатель передает во временное и платное распоряжение Пользователю такие объекты интеллектуальной собственности ______ (товарный знак, торговую марку, технологию и т.д.).",
        "2. Пользователь обязуется использовать полученные права для ______ (производства определенных товаров или оказания услуг).",
        "3. За предоставленные Правообладателем права Пользователь выплачивает вознаграждение в таком порядке: ______.",
        "4. Договор коммерческой концессии подлежит государственной регистрации в установленном законодательством порядке.",
        "5. Правообладатель имеет следующие права и обязательства: ______.",
        "6. У Пользователя имеются такие права и обязанности: ______."
    };

            foreach (string clause in clauses)
            {
                object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Paragraph p = doc.Content.Paragraphs.Add(ref range);
                p.Range.Text = clause;
                p.Range.Font.Name = "Times New Roman";
                p.Range.Font.Size = 12;
                p.Format.SpaceAfter = 6; // Уменьшаем отступ между пунктами
                p.Format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                p.Format.FirstLineIndent = 28; // Красная строка (1.25 см)
                p.Range.InsertParagraphAfter();
            }
        }


        private void AddEmptyParagraph(Document doc, object endOfDoc)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph p = doc.Content.Paragraphs.Add(ref range);
            p.Range.Text = "";
            p.Format.SpaceAfter = 6;
            p.Range.InsertParagraphAfter();
        }

        private void AddEmptyParagraphs(Document doc, object endOfDoc, int count)
        {
            for (int i = 0; i < count; i++)
            {
                AddEmptyParagraph(doc, endOfDoc);
            }
        }
    }
}