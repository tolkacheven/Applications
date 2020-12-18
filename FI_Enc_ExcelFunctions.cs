using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
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
            GetAllCells.Columns.ColumnWidth = 15;

            bool col_width_autofit = (col_width == -1) ? true : false;

            GetAllCells.RowHeight = 50;
            //foreach (Excel.Range row in GetAllCells.Rows)
            //    row.RowHeight = 50;

            if (col_width_autofit == true)
                foreach (Excel.Range column in GetAllCells.Columns)
                    column.AutoFit();
            else
                foreach (Excel.Range column in GetAllCells.Columns)
                    column.ColumnWidth = col_width;

            return 0;
        }



        // Определить индексы колонок (1-ый аргумент - объект типа Worksheet, далее - произвольное количество необходимых заголовков)
        public static Dictionary<string, int> Excel_DefineColumns(Excel.Worksheet ws, params string[] headers)
        {
            int column_index = -1;
            var headers_indexes = new Dictionary <string, int> ();

            Excel.Range ws_cells = ws.UsedRange;

            foreach (string col_title in headers)
            {
                column_index = -1;

                for (int i = 1; i <= ws_cells.Columns.Count; i++)
                {
                    if (ws_cells.Cells[1, i].Value != null)
                    {
                        if (ws_cells.Cells[1, i].Value.ToString() == col_title)
                        {
                            column_index = i;
                            break;
                        }
                    }
                }

                headers_indexes.Add(col_title, column_index);
            }
            
            return headers_indexes;
        }



        // Проверка - находится ли дата в указанном диапазоне?
        public static int IsDateBetweenDates(string required_date, string date_from, string date_to)
        {
            DateTime required_date_converted = ExtractDateFromString(required_date),
                     date_from_converted = ExtractDateFromString(date_from),
                     date_to_converted = ExtractDateFromString(date_to);

            return ((required_date_converted >= date_from_converted) && (required_date_converted <= date_to_converted)) ? 0 : 1;
        }



        // Извлечь дату из строки (и преобразовать в тип DateTime)
        public static DateTime ExtractDateFromString(string input_date)
        {
            var regex = new Regex(@"\b\d{2}\.\d{2}.\d{4}\b");
            DateTime result = new DateTime();

            foreach (Match m in regex.Matches(input_date))
            {
                if (DateTime.TryParseExact(m.Value, "dd.MM.yyyy", null, DateTimeStyles.None, out result))
                    break;
            }

            return result;
        }



        // Функция фильтрации таблицы по указанной колонке и критерию отбора
        public static int Excel_FilterBy(Excel.Worksheet worksheet, int column, string criteria)
        {
            Excel.Range GetAllCells = worksheet.UsedRange;
            GetAllCells.AutoFilter(column, criteria);

            return 0;
        }



        // Функция сведения двух таблиц
        public static int Excel_PivotTable(Excel.Application excel_application, Excel.Worksheet bank_report_ws, string ws1_column, Excel.Worksheet worksheet_2, string ws2_column)
        {
            /*
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            Excel.Range Fruits = Application.get_Range("A1", "B3");
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = Fruits.Find("apples", missing,
                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                missing, missing);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                currentFind.Font.Bold = true;

                currentFind = Fruits.FindNext(currentFind);
            }
            */
            return 0;
        }




        static void Main(string[] args)
        {
            string Message = string.Empty;
            bool Success = false;


            // Конфигурация приложения Эксель
            Excel.Application excel_app = new Excel.Application();
            excel_app.Visible = true;
            excel_app.DisplayAlerts = false;


            /**************************************/
            /*****         АЛЬФА-БАНК         *****/
            /**************************************/
            string alfabank_report_path = "C:\\MagnitTest\\Альфа-Банк\\Альфа-банк.xlsx",
                   alfabank_TID_path = "C:\\MagnitTest\\Альфа-Банк\\Альфа-банк ТИД.xlsx",
                   alfabank_pivot_path = "C:\\MagnitTest\\Альфа-Банк\\Альфа-банк Свод.xlsx",
                   alfabank_ws_name = "Исходные данные",
                   alfabank_TID_ws_name = "Адреса и тиды",
                   alfabank_pivot_source_ws_name = "Исходные данные";


            /**************************************/
            /*****             ВТБ            *****/
            /**************************************/
            string vtb_report_path = "C:\\MagnitTest\\ВТБ\\ВТБ.xlsx",
                   vtb_TID_path = "C:\\MagnitTest\\ВТБ\\ВТБ ТИД.xlsx",
                   vtb_pivot_path = "C:\\MagnitTest\\ВТБ\\ВТБ Свод.xlsx",
                   vtb_ws_name = "Исходник",
                   vtb_TID_ws_name = "Адреса и тиды", 
                   vtb_pivot_source_ws_name = "Исходные данные";


            /**************************************/
            /*****             ГПБ            *****/
            /**************************************/
            string gpb_report_path = "C:\\MagnitTest\\ГПБ\\ГПБ.xlsx",
                   gpb_TID_path = "C:\\MagnitTest\\ГПБ\\ГПБ ТИД.xlsx",
                   gpb_pivot_path = "C:\\MagnitTest\\ГПБ\\ГПБ Свод.xlsx",
                   gpb_ws_name = "Исходные данные",
                   gpb_TID_ws_name = "Единые ТИД",
                   gpb_pivot_source_ws_name = "Исходные данные";


            /**************************************/
            /*****       Входные данные       *****/
            /**************************************/
            string date_from = "01.10.2020", date_to = "01.12.2020";

            string bank = "ВТБ";

            var checking_accounts = new Dictionary <string, string> {{"ГПБ",        "40702810192001012469"},
                                                                     {"ВТБ",        "40702810003300000230"},
                                                                     {"Альфа-банк", "40702810226020005668"}};

            string bank_id = checking_accounts[bank];

            // Отчет робота по конкретному банку
            Excel.Workbook bank_report_wb;
            Excel.Worksheet bank_report_ws;
            string bank_report_ws_name = "",
                   bank_report_path    = "";


            // Справочник
            Excel.Workbook bank_refbook_wb;
            Excel.Worksheet bank_refbook_ws;
            string bank_refbook_ws_name = "",
                   bank_refbook_path    = "";


            // Промежуточная сводная таблица
            Excel.Workbook bank_pivot_wb;
            Excel.Worksheet bank_pivot_ws;
            string bank_pivot_path = "";
            string bank_pivot_ws_name = "";


            switch (bank_id)
            {
                // ГПБ
                case "40702810192001012469":
                    {
                        bank_report_path = gpb_report_path;
                        bank_refbook_path = gpb_TID_path; 
                        bank_report_ws_name = gpb_ws_name; 
                        bank_refbook_ws_name = gpb_TID_ws_name;
                        bank_pivot_path = gpb_pivot_path;
                        bank_pivot_ws_name = gpb_pivot_source_ws_name;
                        break; 
                    }


                // ВТБ
                case "40702810003300000230":
                    {
                        bank_report_path = vtb_report_path;
                        bank_refbook_path = vtb_TID_path;
                        bank_report_ws_name = vtb_ws_name;
                        bank_refbook_ws_name = vtb_TID_ws_name;
                        bank_pivot_path = vtb_pivot_path;
                        bank_pivot_ws_name = vtb_pivot_source_ws_name;
                        break;
                    }

                // Альфа-банк
                case "40702810226020005668":
                    {
                        bank_report_path = alfabank_report_path;
                        bank_refbook_path = alfabank_TID_path;
                        bank_report_ws_name = alfabank_ws_name;
                        bank_refbook_ws_name = alfabank_TID_ws_name;
                        bank_pivot_path = alfabank_pivot_path;
                        bank_pivot_ws_name = alfabank_pivot_source_ws_name;
                        break;
                    }

                default:
                    {
                        throw new Exception("Указан некорректный расчетный счет");
                    }
            }


            // Открыть отчет робота и считать информацию
            bank_report_wb = excel_app.Workbooks.Open(bank_report_path, ReadOnly: true);
            bank_report_ws = bank_report_wb.Worksheets[bank_report_ws_name];
            Excel.Range bank_report_ws_cells = bank_report_ws.UsedRange;

            // Открыть справочник и считать информацию
            bank_refbook_wb = excel_app.Workbooks.Open(bank_refbook_path, ReadOnly: true);
            bank_refbook_ws = bank_refbook_wb.Worksheets[bank_refbook_ws_name];
            Excel.Range bank_refbook_ws_cells = bank_refbook_ws.UsedRange;


            // Открыть промежуточную сводную таблицу и считать информацию
            bank_pivot_wb = excel_app.Workbooks.Open(bank_pivot_path);
            bank_pivot_ws = bank_pivot_wb.Worksheets[bank_pivot_ws_name];
            Excel.Range bank_pivot_ws_cells = bank_pivot_ws.UsedRange;


            var bank_report_headers = Excel_DefineColumns(bank_report_ws, "Подразделение", "Документ", "Дата внесения", "Сумма", "Комментарий");
            var bank_refbook_headers = Excel_DefineColumns(bank_refbook_ws, "Подразделение", "ТИД", "Адрес");
            var bank_pivot_headers = Excel_DefineColumns(bank_pivot_ws, "Подразделение", "Документ", "Период", "Сумма", "Комментарий", "Адрес", "ТИД");

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            Excel.Range pivot_data = bank_pivot_ws.UsedRange;

            Excel.Range pivot_destination = bank_pivot_ws.get_Range("A46", "A46");

            bank_pivot_wb.PivotTableWizard(
                    Excel.XlPivotTableSourceType.xlDatabase,
                    pivot_data,
                    pivot_destination,
                    "Исходные данные",
                    true,
                    true,
                    true,
                    true,
                    Type.Missing,
                    Type.Missing,
                    false,
                    false,
                    Excel.XlOrder.xlDownThenOver,
                    0,
                    Type.Missing,
                    Type.Missing
            );

            // Set variables used to manipulate the Pivot Table.
            Excel.PivotTable pivot_table = (Excel.PivotTable) bank_pivot_ws.PivotTables("Исходные данные");

            Excel.PivotField Y = ((Excel.PivotField) pivot_table.PivotFields("Период"));
            Excel.PivotField M = ((Excel.PivotField) pivot_table.PivotFields("Подразделение"));
            Excel.PivotField sum_of_doc = ((Excel.PivotField) pivot_table.PivotFields("Сумма"));

            Y.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            M.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            sum_of_doc.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
            sum_of_doc.Function = Excel.XlConsolidationFunction.xlSum;


            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            /*
            foreach (var item in bank_report_headers)
                Console.WriteLine("Report " + item.Key + " -> " + item.Value);

            foreach (var item in bank_refbook_headers)
                Console.WriteLine("Refbook: " + item.Key + " -> " + item.Value);

            foreach (var item in bank_pivot_headers)
                Console.WriteLine("Pivot: " + item.Key + " -> " + item.Value);



            //try
            //{
                if (excel_app == null)
                    throw new Exception("Excel could not be started");

                string report_primary_key = "", pivot_primary_key = "", current_date = "";
                bool document_exists = false;
                int first_empty_row = bank_pivot_ws_cells.Rows.Count + 1;


                for (int i = 2; i <= bank_report_ws_cells.Rows.Count; i++)
                {
                    document_exists = false;
                    report_primary_key = bank_report_ws.Cells[i, bank_report_headers["Документ"]].Value.ToString();
                    current_date = bank_report_ws_cells.Cells[i, bank_report_headers["Дата внесения"]].Value.ToString();
                    
                    if (IsDateBetweenDates(current_date, date_from, date_to) != 0)
                        continue;

                    for (int j = 2; j <= bank_pivot_ws_cells.Rows.Count; j++)
                    {
                        pivot_primary_key = bank_pivot_ws.Cells[j, bank_pivot_headers["Документ"]].Value.ToString();
                    
                        if (report_primary_key == pivot_primary_key) 
                        {
                            document_exists = true;
                            break;
                        }                    
                    }

                    if (document_exists == false)
                    {
                        bank_pivot_ws_cells.Cells[first_empty_row, bank_pivot_headers["Период"]].Value = bank_report_ws_cells.Cells[i, bank_report_headers["Дата внесения"]].Value.ToString();
                        bank_pivot_ws_cells.Cells[first_empty_row, bank_pivot_headers["Документ"]].Value = bank_report_ws_cells.Cells[i, bank_report_headers["Документ"]].Value.ToString();
                        bank_pivot_ws_cells.Cells[first_empty_row, bank_pivot_headers["Сумма"]].Value = bank_report_ws_cells.Cells[i, bank_report_headers["Сумма"]].Value.ToString();
                        bank_pivot_ws_cells.Cells[first_empty_row, bank_pivot_headers["Подразделение"]].Value = bank_report_ws_cells.Cells[i, bank_report_headers["Подразделение"]].Value.ToString();

                        if (bank_report_ws_cells.Cells[i, bank_report_headers["Комментарий"]].Value != null)
                            bank_pivot_ws_cells.Cells[first_empty_row, bank_pivot_headers["Комментарий"]].Value = bank_report_ws_cells.Cells[i, bank_report_headers["Комментарий"]].Value.ToString();

                        first_empty_row++;
                    }
                }

                bank_pivot_wb.Save();

                bank_pivot_ws_cells = bank_pivot_ws.UsedRange;

                for (int i = 2; i <= bank_pivot_ws_cells.Rows.Count; i++) {

                    if (bank_pivot_ws_cells.Cells[i, bank_pivot_headers["ТИД"]].Value != null)
                        continue;

                    for (int j = 2; j <= bank_refbook_ws_cells.Rows.Count; j++)
                    {
                        if (bank_pivot_ws_cells.Cells[i, bank_pivot_headers["Подразделение"]].Value.ToString() == bank_refbook_ws_cells.Cells[j, bank_refbook_headers["Подразделение"]].Value.ToString())
                        {
                            bank_pivot_ws_cells.Cells[i, bank_pivot_headers["ТИД"]].Value = bank_refbook_ws_cells.Cells[j, bank_refbook_headers["ТИД"]].Value.ToString();
                            if (bank_refbook_headers["Адрес"] > 0)
                                bank_pivot_ws_cells.Cells[i, bank_pivot_headers["Адрес"]].Value = bank_refbook_ws_cells.Cells[j, bank_refbook_headers["Адрес"]].Value.ToString();
                            break;
                        }
                    }
                }            
                bank_pivot_wb.Save();
            */
            //}
            //catch (Exception e) { Message = Convert.ToString(e); }
            //finally { bank_report_wb.Close(false); bank_refbook_wb.Close(false); bank_pivot_wb.Close(SaveChanges: true, Filename: bank_pivot_path); }
        }
    }
}
