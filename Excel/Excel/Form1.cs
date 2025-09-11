using System.Reflection;
using ExcelApp = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using System;

namespace VasilevPracticeExcel
{
    public partial class Form1 : Form
    {
        string[] csvData;

        public Form1()
        {
            InitializeComponent();

            string csvFilePath = "defaulData.csv";
            if (!File.Exists(csvFilePath))
            {
                CreateDefaultCsvFile(csvFilePath);
            }

            csvData = File.ReadAllText(csvFilePath, System.Text.Encoding.UTF8).Split(',');

            for (int i = 0; i < csvData.Length; i += 9)
            {
                if (i + 8 < csvData.Length)
                {
                    PreviewDGV.Rows.Add(
                        csvData[i], csvData[i + 1], csvData[i + 2],
                        csvData[i + 3], csvData[i + 4], csvData[i + 5],
                        csvData[i + 6], csvData[i + 7], csvData[i + 8]
                    );
                }
            }
        }

        private void CreateDefaultCsvFile(string filePath)
        {
            // Создаем пример данных для материального аккаунта
            string[] defaultData = {
                "2025-01-15", "ТН-001", "1", "Поставщик ООО 'ВООВОВИО'", "Сталь", "500", "0", "500", "Иванов 15.01.2025",
                "2025-01-16", "РН-001", "2", "Цех механической обработки", "Сталь листовая", "0", "200", "300", "Петров 16.01.2025",
                "2025-01-17", "ТН-002", "3", "Поставщик ИП луьщьдс", "Болты М", "1000", "0", "1000", "Федоров 17.01.2025",
                "2025-01-18", "РН-002", "4", "Сборочный цех", "Болты ", "0", "300", "700", "Козлов 18.01.2025"
            };

            File.WriteAllText(filePath, string.Join(",", defaultData));
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelApp.Application app = new ExcelApp.Application();
                app.Visible = false;
                ExcelApp.Workbook wb = app.Workbooks.Add(Missing.Value);
                ExcelApp.Worksheet ws = (ExcelApp.Worksheet)wb.Sheets[1];
                ws.Activate();

                // Создаем заголовок с датой
                ws.Range["A1", "I1"].Merge();
                ws.Cells[1, 1] = "Дата записи: " + DateTime.Now.ToString("dd.MM.yyyy");
                ws.Range["A1", "I1"].Font.Bold = true;
                ws.Range["A1", "I1"].HorizontalAlignment = ExcelApp.XlHAlign.xlHAlignCenter;

                // Основная шапка таблицы
                ws.Range["A2", "B2"].Merge();
                ws.Cells[2, 1] = "Номер документа";

                ws.Cells[2, 3] = "От кого получено или кому отправлено";
                ws.Range["C2", "C3"].Merge();

                ws.Cells[2, 4] = "Учетная единица выпуска продукции (работ, услуг)";
                ws.Range["D2", "D3"].Merge();

                ws.Range["E2", "G2"].Merge();
                ws.Cells[2, 5] = "Движение материалов";

                ws.Range["H2", "I2"].Merge();
                ws.Cells[2, 8] = "Подпись, дата";
                ws.Range["H2", "I3"].Merge();

                // Подзаголовки
                ws.Cells[3, 1] = "по порядку";
                ws.Cells[3, 2] = "документа";
                ws.Cells[3, 5] = "Приход";
                ws.Cells[3, 6] = "Расход";
                ws.Cells[3, 7] = "Остаток";

                // Нумерация столбцов
                for (int i = 1; i <= 9; i++)
                {
                    ws.Cells[4, i] = i.ToString();
                }

                // Настройка ширины столбцов
                ws.Columns[1].ColumnWidth = 8;
                ws.Columns[2].ColumnWidth = 12;
                ws.Columns[3].ColumnWidth = 30;
                ws.Columns[4].ColumnWidth = 25;
                ws.Columns[5].ColumnWidth = 10;
                ws.Columns[6].ColumnWidth = 10;
                ws.Columns[7].ColumnWidth = 10;
                ws.Columns[8].ColumnWidth = 15;
                ws.Columns[9].ColumnWidth = 12;

                // Заполнение данными
                for (int i = 0; i < csvData.Length; i++)
                {
                    int row = 5 + i / 9;
                    int col = 1 + i % 9;
                    ws.Cells[row, col] = csvData[i];
                }

                // Форматирование таблицы
                ExcelApp.Range tableRange = ws.Range["A2", $"I{4 + csvData.Length / 9}"];
                tableRange.Borders.LineStyle = ExcelApp.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = ExcelApp.XlBorderWeight.xlThin;

                // Выравнивание
                tableRange.HorizontalAlignment = ExcelApp.XlHAlign.xlHAlignCenter;
                tableRange.VerticalAlignment = ExcelApp.XlVAlign.xlVAlignCenter;

                // Жирный шрифт для заголовков
                ws.Range["A2", "I4"].Font.Bold = true;

                // Автоподбор высоты строк
                ws.Rows.AutoFit();

                app.Visible = true;

                // Сохранение файла
                string fileName = $"Material_Account_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                wb.SaveAs(fileName);

                MessageBox.Show($"Файл успешно сохранен: {fileName}", "Успех",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}