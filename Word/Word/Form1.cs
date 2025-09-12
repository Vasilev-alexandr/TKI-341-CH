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

namespace WordAutomation
{
    public partial class Form1 : Form
    {
        private string configFile = "config.txt";
        private string authorName = "Васильев А.С.";
        private int leftMarginOffset = 0;
        private Microsoft.Office.Interop.Word.Application wordApp;

        public Form1()
        {
            InitializeComponent();
            SetupForm();
            LoadConfig();
        }

        private void SetupForm()
        {
            this.Text = $"Задание №5 выполнил: Васильев А.С.; Номер варианта: 3; Дата выполнения: {DateTime.Now:dd/MM/yyyy}";

            cmbDocumentType.Items.AddRange(new string[] { "отчёт", "реферат", "эссе", "курсовой проект", "курсовая работа", "доклад", "домашнее задание" });
            cmbWorkType.Items.AddRange(new string[] { "лабораторная работа", "практическая работа", "индивидуальное задание", "учебная практика", "производственная практика", "преддипломная практика" });
            cmbWorkNumber.Items.AddRange(Enumerable.Range(1, 10).Select(i => i.ToString()).ToArray());

            cmbDocumentType.SelectedIndex = 0;
            cmbWorkType.SelectedIndex = 0;
            cmbWorkNumber.SelectedIndex = 0;

            trackBarLeftMargin.Minimum = 0;
            trackBarLeftMargin.Maximum = 200;
            trackBarLeftMargin.Value = 0;
            trackBarLeftMargin.TickFrequency = 10;
            lblMarginValue.Text = "0";
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
                    $"{cmbDocumentType.Text};Times New Roman;16",
                    $"{txtLink1.Text};Times New Roman;14",
                    $"{cmbWorkType.Text};Times New Roman;14",
                    $"{txtLink2.Text};Times New Roman;14",
                    $"{txtDiscipline.Text};Times New Roman;14",
                    $"{txtLink3.Text};Times New Roman;14",
                    $"{txtGroup.Text};Times New Roman;14",
                    $"{txtAuthor.Text};Times New Roman;14",
                    $"{txtChecker.Text};Times New Roman;14",
                    $"{txtCityYear.Text};Times New Roman;14"
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
                leftMarginOffset = trackBarLeftMargin.Value;
                CreateWordDocument();
                MessageBox.Show("Документ создан успешно!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка создания документа: {ex.Message}");
            }
        }

        private void trackBarLeftMargin_Scroll(object sender, EventArgs e)
        {
            leftMarginOffset = trackBarLeftMargin.Value;
            lblMarginValue.Text = leftMarginOffset.ToString();
        }

        private void ConfigureDocumentStyles(Document doc)
        {
            doc.Styles[WdBuiltinStyle.wdStyleNormal].Font.Name = "Times New Roman";
            doc.Styles[WdBuiltinStyle.wdStyleNormal].Font.Size = 14;
            doc.Styles[WdBuiltinStyle.wdStyleNormal].ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            doc.Styles[WdBuiltinStyle.wdStyleNormal].ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            doc.Styles[WdBuiltinStyle.wdStyleNormal].ParagraphFormat.SpaceAfter = 0;
        }

        private void CreateWordDocument()
        {
            object missing = Missing.Value;
            object endOfDoc = "\\endofdoc";

            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;

            try
            {
                Document doc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                doc.Activate();

                ConfigureDocumentStyles(doc);

                AddTitlePage(doc, endOfDoc);

                object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Range breakRange = (Range)range;
                breakRange.InsertBreak(Type: WdBreakType.wdPageBreak);

                AddContractContent(doc, endOfDoc);

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

        private void AddTitlePage(Document doc, object endOfDoc)
        {
            doc.PageSetup.LeftMargin = 0;
            doc.PageSetup.RightMargin = 0;

            AddCenteredText(doc, endOfDoc, "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ", 14, false);
            AddCenteredText(doc, endOfDoc, "ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ", 14, false);
            AddCenteredText(doc, endOfDoc, "«РОССИЙСКИЙ УНИВЕРСИТЕТ ТРАНСПОРТА»", 14, false);
            AddCenteredText(doc, endOfDoc, "(РУТ (МИИТ))", 14, false);
            AddEmptyParagraph(doc, endOfDoc);
            AddEmptyParagraph(doc, endOfDoc);
            AddEmptyParagraph(doc, endOfDoc);
            AddCenteredText(doc, endOfDoc, "Институт транспортной техники и систем управления", 14, false);
            AddCenteredText(doc, endOfDoc, "Кафедra «Управление и защита информации»", 14, false);
            AddEmptyParagraph(doc, endOfDoc);
            AddEmptyParagraph(doc, endOfDoc);
            AddEmptyParagraph(doc, endOfDoc);
            AddEmptyParagraph(doc, endOfDoc);
            AddCenteredText(doc, endOfDoc, "ОТЧЁТ", 14, false);
            AddCenteredText(doc, endOfDoc, "по практической работе", 14, false);
            AddCenteredText(doc, endOfDoc, "«Договор коммерческой концессии»", 14, true);
            AddCenteredText(doc, endOfDoc, "По дисциплине «Алгоритмизация и технологии программирования»", 14, false);

            for (int i = 0; i < 7; i++)
            {
                AddEmptyParagraph(doc, endOfDoc);
            }

            AddRightAlignedText(doc, endOfDoc, "Выполнил: ст. гр. ТКИ – 341", 14, false);

            AddRightAlignedText(doc, endOfDoc, "Васильев А. С.", 14, false);

            AddRightAlignedText(doc, endOfDoc, "Проверил: доц. к.т.н. Сафронов А. И.", 14, false);

            for (int i = 0; i < 10; i++)
            {
                AddEmptyParagraph(doc, endOfDoc);
            }

            AddCenteredText(doc, endOfDoc, "Москва 2025", 14, false);
        }

        private void AddContractContent(Document doc, object endOfDoc)
        {
            doc.PageSetup.LeftMargin = 0;
            doc.PageSetup.RightMargin = 0;

            doc.Styles[WdBuiltinStyle.wdStyleNormal].Font.Size = 11;

            AddCenteredTextWithSpacing(doc, endOfDoc, "ДОГОВОР", 11, true, 6, 0);

            AddCenteredTextWithSpacing(doc, endOfDoc, "коммерческой концессии", 11, false, 6, 0);

            AddEmptyParagraph(doc, endOfDoc);

            AddCityDateLine(doc, endOfDoc);

            AddEmptyParagraphWithSpacing(doc, endOfDoc, 12, 6);

            AddContractParties(doc, endOfDoc);

            AddContractText(doc, endOfDoc);
        }

        private void AddCityDateSameLine(Document doc, object endOfDoc)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph paragraph = doc.Content.Paragraphs.Add(ref range);

            Table table = doc.Tables.Add((Range)range, 1, 2);
            table.Borders.Enable = 0;

            table.Cell(1, 1).Range.Text = "г. _______";
            table.Cell(1, 1).Range.Font.Name = "Times New Roman";
            table.Cell(1, 1).Range.Font.Size = 11;
            table.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

            table.Cell(1, 2).Range.Text = "«___» __ 20__ г.";
            table.Cell(1, 2).Range.Font.Name = "Times New Roman";
            table.Cell(1, 2).Range.Font.Size = 11;
            table.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;

            table.Columns[1].PreferredWidth = wordApp.CentimetersToPoints(8);
            table.Columns[2].PreferredWidth = wordApp.CentimetersToPoints(8);

            table.Range.InsertParagraphAfter();
        }

        private void AddContractPartiesCorrected(Document doc, object endOfDoc)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph owner = doc.Content.Paragraphs.Add(ref range);
            owner.Range.Text = "«Правообладатель» __________________________, в лице ____________, действующего на основании _____________________________, и";
            owner.Range.Font.Name = "Times New Roman";
            owner.Range.Font.Size = 11;
            owner.Format.SpaceBefore = 12;
            owner.Format.SpaceAfter = 6;
            owner.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            owner.Range.InsertParagraphAfter();

            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph user = doc.Content.Paragraphs.Add(ref range);
            user.Range.Text = "«Пользователь» _______________________, в лице _______________, действующего на основании ______________________, подписали этот договор коммерческой концессии на таких условиях.";
            user.Range.Font.Name = "Times New Roman";
            user.Range.Font.Size = 11;
            user.Format.SpaceBefore = 0;
            user.Format.SpaceAfter = 12;
            user.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            user.Range.InsertParagraphAfter();
        }

        private void AddCenteredText(Document doc, object endOfDoc, string text, int fontSize, bool bold)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph paragraph = doc.Content.Paragraphs.Add(ref range);
            paragraph.Range.Text = text;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Size = fontSize;
            paragraph.Range.Font.Bold = bold ? 1 : 0;
            paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Format.SpaceAfter = 0;
            paragraph.Range.InsertParagraphAfter();
        }

        private void AddCenteredTextWithSpacing(Document doc, object endOfDoc, string text, int fontSize, bool bold, int spaceBefore, int spaceAfter)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph paragraph = doc.Content.Paragraphs.Add(ref range);
            paragraph.Range.Text = text;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Size = fontSize;
            paragraph.Range.Font.Bold = bold ? 1 : 0;
            paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.Format.SpaceBefore = spaceBefore;
            paragraph.Format.SpaceAfter = spaceAfter;
            paragraph.Range.InsertParagraphAfter();
        }

        private void AddRightAlignedText(Document doc, object endOfDoc, string text, int fontSize, bool bold)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph paragraph = doc.Content.Paragraphs.Add(ref range);
            paragraph.Range.Text = text;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Size = fontSize;
            paragraph.Range.Font.Bold = bold ? 1 : 0;
            paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            paragraph.Format.SpaceAfter = 0;
            paragraph.Range.InsertParagraphAfter();
        }

        private void AddEmptyParagraph(Document doc, object endOfDoc)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph p = doc.Content.Paragraphs.Add(ref range);
            p.Range.Text = "";
            p.Format.SpaceAfter = 0;
            p.Range.InsertParagraphAfter();
        }

        private void AddEmptyParagraphWithSpacing(Document doc, object endOfDoc, int spaceBefore, int spaceAfter)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph p = doc.Content.Paragraphs.Add(ref range);
            p.Range.Text = "";
            p.Format.SpaceBefore = spaceBefore;
            p.Format.SpaceAfter = spaceAfter;
            p.Range.InsertParagraphAfter();
        }

        private void AddCityDateLine(Document doc, object endOfDoc)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph paragraph = doc.Content.Paragraphs.Add(ref range);

            paragraph.Range.Text = "г. _______                                                                                              «____» _______ 20__ г.";
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Size = 11;
            paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.Format.SpaceAfter = 0;
            paragraph.Format.FirstLineIndent = 0;
            paragraph.Range.InsertParagraphAfter();
        }

        private void AddContractParties(Document doc, object endOfDoc)
        {
            object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph line6 = doc.Content.Paragraphs.Add(ref range);
            line6.Range.Text = "«Правообладатель» __________________________, в лице ____________, действующего";
            line6.Range.Font.Name = "Times New Roman";
            line6.Range.Font.Size = 11;
            line6.Format.SpaceBefore = 12;
            line6.Format.SpaceAfter = 6;
            line6.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            line6.Format.FirstLineIndent = 0;
            line6.Range.InsertParagraphAfter();

            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph line7 = doc.Content.Paragraphs.Add(ref range);
            line7.Range.Text = "на основании _____________________________, и";
            line7.Range.Font.Name = "Times New Roman";
            line7.Range.Font.Size = 11;
            line7.Format.SpaceBefore = 0;
            line7.Format.SpaceAfter = 6;
            line7.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            line7.Format.FirstLineIndent = 0;
            line7.Range.InsertParagraphAfter();

            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph line8 = doc.Content.Paragraphs.Add(ref range);
            line8.Range.Text = "«Пользователь» _______________________, в лице _______________, действующего на";
            line8.Range.Font.Name = "Times New Roman";
            line8.Range.Font.Size = 11;
            line8.Format.SpaceBefore = 0;
            line8.Format.SpaceAfter = 6;
            line8.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            line8.Format.FirstLineIndent = 0;
            line8.Range.InsertParagraphAfter();

            range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
            Paragraph line9 = doc.Content.Paragraphs.Add(ref range);
            line9.Range.Text = "основании ______________________, подписали этот договор коммерческой концессии на таких условиях.";
            line9.Range.Font.Name = "Times New Roman";
            line9.Range.Font.Size = 11;
            line9.Format.SpaceAfter = 0;
            line9.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            line9.Format.FirstLineIndent = 0;
            line9.Range.InsertParagraphAfter();
        }

        private void AddContractText(Document doc, object endOfDoc)
        {
            string[] clauses = {
        "1. Правообладатель передает во временное и платное распоряжение Пользователю такие объекты",
        "интеллектуальной собственности ___________________ (товарный знак, торговую марку,",
        "технологию и т.д.).",
        "2. Пользователь обязуется использовать полученные права для ______________________",
        "_________________________________ (производства определенных товаров или оказания услуг).",
        "3. За предоставленные Правообладателем права Пользователь выплачивает вознаграждение в",
        "таком порядке: _________________________________________________.",
        "4. Договор коммерческой концессии подлежит государственной регистрации в установленном",
        "законодательством порядке.",
        "5. Правообладатель имеет следующие права и обязательства: ________________________",
        "________________________________________________________________________________.",
        "6. У Пользователя имеются такие права и обязанности: _____________________________",
        "________________________________________________________________________________."
    };

            foreach (string clause in clauses)
            {
                object range = doc.Bookmarks.get_Item(ref endOfDoc).Range;
                Paragraph p = doc.Content.Paragraphs.Add(ref range);
                p.Range.Text = clause;
                p.Range.Font.Name = "Times New Roman";
                p.Range.Font.Size = 11;
                p.Format.SpaceAfter = 0;
                p.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                p.Format.FirstLineIndent = 0;
                p.Range.InsertParagraphAfter();
            }
        }
    }
}
