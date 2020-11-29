using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace FI_Enc_TablesConcatenation
{
    class Program
    {
        /* Форматирование Эксель файла:
         * 1-ый аргумент - объект класса Worksheet;
         * 2-ой аргумент - высота каждой строки;
         * 3-ий аргумент - ширина каждой колонки (по умолчанию Autofit).
         */
        public static int Excel_Autofit(Excel.Worksheet worksheet, int row_height = 50, int col_width = -1)
        {
            Excel.Range GetAllCells = worksheet.UsedRange;

            bool col_width_autofit = (col_width == -1) ? true : false;

            foreach (Excel.Range row in GetAllCells.Rows)
                row.RowHeight = 50;

            if (col_width_autofit == true)
                foreach (Excel.Range column in GetAllCells.Columns)
                    column.AutoFit();
            else
                foreach (Excel.Range column in GetAllCells.Columns)
                    column.ColumnWidth = col_width;

            return 0;
        }


        // Функция фильтрации таблицы по указанной колонке и критерию отбора
        public static int Excel_FilterBy(Excel.Worksheet worksheet, int column, string criteria)
        {
            Excel.Range GetAllCells = worksheet.UsedRange;
            GetAllCells.AutoFilter(column, criteria);

            return 0;
        }


        // Функция сведения двух таблиц
        public static int Excel_PivotTable(Excel.Application excel_application, Excel.Worksheet worksheet_1, string ws1_column, Excel.Worksheet worksheet_2, string ws2_column)
        {

            return 0;
        }



        static void Main(string[] args)
        {
            string Message = string.Empty;
            bool Success = false;


            // Конфигурация приложения Эксель
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;


            // Первая таблица (Магнит)
            Excel.Workbook workbook_1;
            Excel.Worksheet worksheet_1;
            string worksheet_name_1 = "Sheet1";
            string table1_filepath = "C:\\MagnitTest\\Report17.11.2020.xlsx";


            // Промежуточная сводная таблица
            Excel.Workbook workbook_pivot;
            Excel.Worksheet worksheet_pivot;
            string worksheet_name_pivot = "Альфа-Банк";
            string pivot_filepath = "C:\\MagnitTest\\PivotTable.xlsx";

            
            // Справочник
            Excel.Workbook workbook_refbook;
            Excel.Worksheet worksheet_refbook;
            string worksheet_name_refbook = "Вкладка№2";
            string refbook_filepath = "C:\\MagnitTest\\Справочник1.xlsx";


            // Открыть отчет робота и считать информацию
            workbook_1 = excelApp.Workbooks.Open(table1_filepath, ReadOnly: true);
            worksheet_1 = workbook_1.Worksheets[worksheet_name_1];
            Excel.Range worksheet_1_cells = worksheet_1.UsedRange;


            // Открыть справочник и считать информацию
            workbook_refbook = excelApp.Workbooks.Open(refbook_filepath, ReadOnly: true);
            worksheet_refbook = workbook_refbook.Worksheets[worksheet_name_refbook];
            Excel.Range worksheet_refbook_cells = worksheet_refbook.UsedRange;


            // Открыть промежуточную сводную таблицу и считать информацию
            workbook_pivot = excelApp.Workbooks.Open(pivot_filepath);
            worksheet_pivot = workbook_pivot.Worksheets[worksheet_name_pivot];
            Excel.Range worksheet_pivot_cells = worksheet_pivot.UsedRange;


            try
            {
                if (excelApp == null)
                    throw new Exception("Excel could not be started");

                string report_primary_key = "", pivot_primary_key = "";
                bool document_exists = false;
                int first_empty_row = worksheet_pivot_cells.Rows.Count + 1;


                for (int i = 2; i <= worksheet_1_cells.Rows.Count; i++)
                {
                    document_exists = false;
                    report_primary_key = worksheet_1.Cells[i, 2].Value.ToString();

                    for (int j = 2; j <= worksheet_pivot_cells.Rows.Count; j++)
                    {
                        pivot_primary_key = worksheet_pivot.Cells[j, 2].Value.ToString();
                    
                        if (report_primary_key == pivot_primary_key) 
                        {
                            document_exists = true;
                            break;
                        }                    
                    }

                    if (document_exists == false)
                    {
                        worksheet_pivot.Cells[first_empty_row, 1].Value = worksheet_1.Cells[i, 1].Value.ToString();
                        worksheet_pivot.Cells[first_empty_row, 2].Value = worksheet_1.Cells[i, 2].Value.ToString();
                        worksheet_pivot.Cells[first_empty_row, 3].Value = worksheet_1.Cells[i, 6].Value.ToString();
                        worksheet_pivot.Cells[first_empty_row, 4].Value = worksheet_1.Cells[i, 8].Value.ToString();
                        worksheet_pivot.Cells[first_empty_row, 5].Value = worksheet_1.Cells[i, 20].Value.ToString();
                        worksheet_pivot.Cells[first_empty_row, 6].Value = worksheet_1.Cells[i, 24].Value.ToString();
                        worksheet_pivot.Cells[first_empty_row, 7].Value = worksheet_1.Cells[i, 29].Value.ToString();
                        worksheet_pivot.Cells[first_empty_row, 8].Value = worksheet_1.Cells[i, 30].Value.ToString();

                        first_empty_row++;
                    }
                }
            

                workbook_pivot.Save();

                worksheet_pivot_cells = worksheet_pivot.UsedRange;
            
                for (int i = 2; i <= worksheet_pivot_cells.Rows.Count; i++) {

                    if (worksheet_pivot.Cells[i, 9].Value != null)
                        continue;

                    for (int j = 2; j <= worksheet_refbook_cells.Rows.Count; j++)
                    {
                        if (worksheet_pivot.Cells[i, 6].Value.ToString() == worksheet_refbook.Cells[j, 1].Value.ToString())
                        {
                            worksheet_pivot.Cells[i, 8].Value = worksheet_refbook.Cells[j, 2].Value.ToString();
                            worksheet_pivot.Cells[i, 9].Value = worksheet_refbook.Cells[j, 4].Value.ToString();
                            break;
                        }
                    }
                }
            
                workbook_pivot.Save();
            }
            catch (Exception e) { Message = Convert.ToString(e); }
            finally { workbook_1.Close(false); workbook_refbook.Close(false); workbook_pivot.Close(SaveChanges: true, Filename: pivot_filepath); }
        }
    }
}
