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

                ws.Range["A1", "A2"].Merge(); 
                ws.Cells[1, 1] = "Дата записи";

                ws.Range["B1", "C1"].Merge(); 
                ws.Cells[1, 2] = "Номер";

                ws.Range["D1", "D2"].Merge(); 
                ws.Cells[1, 4] = "От кого получено или кому отпущено";

                ws.Range["E1", "E2"].Merge();
                ws.Cells[1, 5] = "Учетная единица выпуска продукции (работ, услуг)";

                ws.Range["F1", "F2"].Merge(); 
                ws.Cells[1, 6] = "Приход";

                ws.Range["G1", "G2"].Merge();
                ws.Cells[1, 7] = "Расход";

                ws.Range["H1", "H2"].Merge(); 
                ws.Cells[1, 8] = "Остаток";

                ws.Range["I1", "I2"].Merge();
                ws.Cells[1, 9] = "Подпись, дата";

                ws.Cells[2, 2] = "документа";
                ws.Cells[2, 3] = "по порядку"; 

                ws.Cells[3, 1] = "1";
                ws.Cells[3, 2] = "2";
                ws.Cells[3, 3] = "3";
                ws.Cells[3, 4] = "4"; 
                ws.Cells[3, 5] = "5"; 
                ws.Cells[3, 6] = "6"; 
                ws.Cells[3, 7] = "7";
                ws.Cells[3, 8] = "8"; 
                ws.Cells[3, 9] = "9"; 

                ws.Columns[1].ColumnWidth = 12; 
                ws.Columns[2].ColumnWidth = 10;
                ws.Columns[3].ColumnWidth = 8;  
                ws.Columns[4].ColumnWidth = 30; 
                ws.Columns[5].ColumnWidth = 25;
                ws.Columns[6].ColumnWidth = 10; 
                ws.Columns[7].ColumnWidth = 10; 
                ws.Columns[8].ColumnWidth = 10; 
                ws.Columns[9].ColumnWidth = 15; 

                for (int i = 0; i < csvData.Length; i++)
                {
                    int row = 4 + i / 9; 
                    int col = 1 + i % 9;
                    ws.Cells[row, col] = csvData[i];
                }

                ExcelApp.Range headerRange = ws.Range["A1", "I3"];
                headerRange.Font.Bold = true;
                headerRange.HorizontalAlignment = ExcelApp.XlHAlign.xlHAlignCenter;
                headerRange.VerticalAlignment = ExcelApp.XlVAlign.xlVAlignCenter;

                int dataRows = (csvData.Length + 8) / 9;
                ExcelApp.Range tableRange = ws.Range["A1", $"I{3 + dataRows}"];
                tableRange.Borders.LineStyle = ExcelApp.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = ExcelApp.XlBorderWeight.xlThin;

                ws.Rows.AutoFit();

                ws.Range["D1", "E2"].WrapText = true;

                app.Visible = true;

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
