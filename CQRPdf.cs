﻿using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing.Imaging;
using System.Drawing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Excel = Microsoft.Office.Interop.Excel;
using MessagingToolkit.QRCode.Codec;
using MessagingToolkit.QRCode.Codec.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QRPDF
{
    public class CQRPdf
    {
        /*****************************************************************************************************************
         *                                                       ОПИСАНИЕ
         * ***************************************************************************************************************
         * 
         * ### Поставленная задача: написать библиотеку, позволяющую:
         *  >  Генерировать по входным данным QR-коды;
         *  >  Размещать созданные QR-коды в переданных PDF-файлах;
         *  >  Распознавать QR-коды и возвращать их значения под уникальными ID;
         *  >  Поточно обрабатывать переданные файлы.
         *  
         *  
         * ****************************************************************************************************************
         *                                                    ДОПОЛНИТЕЛЬНО
         * ****************************************************************************************************************
         * 
         *  >  Позицию QR-кодов в переданных файлах считаем фиксированной (для уменьшения процента погреш-
         *     ности);
         *  >  Переданные файлы также считаем типовыми (зависит от флага, с которым запущенна программа, а
         *     также меняется в режиме конфигурации);
         *  >  При установке QR-кода создается копия файла;
         *  >  Директория по умолчанию для хранения QRFile(*): 
         *                          C:\Secured Directory\Signification.txt;
         *  >  Директория по умолчанию для хранения расшифрованных QR-кодов и их ID (Excel таблица):
         *                          C:\Secured Directory\QRSpreadsheet.xslx;
         *                          
         * (*) QRFile - файл с информацией, которая должна быть зашифрована в QR-код.
         * 
         * 
         * *****************************************************************************************************************
         *                                                     ПОЛЯ И ФЛАГИ
         * *****************************************************************************************************************
         * 
         *  >  InputFile (директория, string, private):
         *     Директория входного файла, который в зависимости от режима либо будет либо расшифрован, либо зашифрован.
         *     В случае флага Multithread Mode (-M, поточная обработка) берет все файлы из указанной директории.
         *     > Аксессор: InputFilePath (read and write)
         *     
         *     
         *  >  DecodedQRCodesSpreadsheet (директория, string, private):
         *     Файл, в который сохраняются одиночно/поточно обработанные файлы;
         *     > Аксессор: DecodedQRSpreadsheetPath (read only)
         *     
         *     
         *  >  QRInfoFile (директория, string, private):
         *     Директория, в которой хранится информация для шифрования в QR-код;
         *     > Аксессор: QRInfoFilePath (read only)
         *     
         *     
         *  >  ConfigurationsFile (директория, string, private):
         *     Директория, в которой хранится файл конфигурации программы;
         *     > Аксессор: ConfigurationsFilePath (read only)
         *     
         *     
         *  >  Configurations (массив, bool, private):
         *     Массив булевого типа, отвечающий за конфигурацию программы. Каждый элемент обозначает конкретный флаг,
         *     с которым была запущена программа. А именно:
         *     >  Configurations[0] - Stamp Mode       (S) - режим установки QR-кода;
         *     >  Configurations[1] - Decode Mode      (D) - режим расшифровки QR-кода;
         *     >  Configurations[2] - Config Mode      (C) - режим конфигурации программы;
         *     >  Configurations[3] - Multithread Mode (M) - режим поточной обработки файлов;
         *     >  Configurations[4] - Help Mode        (H) - режим "помощи", для отображения всех флагов.
         *
         * 
         * ******************************************************************************************************************
         *                                               ПРОГРАММНЫЙ ИНТЕРФЕЙС
         * ******************************************************************************************************************
         * 
         *  >  public CQRPdf()          - Конструктор по умолчанию. Создает экземпляр класса с предустановленными значениями;
         *  
         *  
         *  >  public CQRPdf(char flag) - Конструктор с одним параметром. Значения флагов:
         *                                > S - Stamp Mode;
         *                                > D - Decode Mode;
         *                                > C - Config Mode;
         *                                > M - Multithread Mode;
         *                                > H - Help Mode.
         *                                
         *                                
         *  >  private CQRPdf(string,   - Конструктор, вызываемый только режимом конфигурации, все три параметра обозначают три директории:
         *                    string,     InputFilePath, DecodedQRCodesSpreadsheet, QRInfoFilePath.
         *                    string)
         *                    
         *                    
         *  >  public void PDFStampQRCode(iTextSharp.text.Image QR,
         *                                int x_offset = 0,   
         *                                int y_offset = 750)
         *                                
         *                              - Функция установки QR-кода. Принимает три параметра:
         *                                > iTextSharp.text.Image QR - преобразованный QR-код (из типа BarcodeQRCode в Image);
         *                                > x_offset - опциональный параметр (по умолчанию   0) - отступ QR-кода по оси x;
         *                                > y_offset - опциональный параметр (по умолчанию 750) - отступ QR-кода по оси y.
         *                                
         *                                
         *  >  public iTextSharp.text.Image QRGenerate(string QRFileContent)
         *                              - Функция, генерирующая QR-код по содержимому файла. Принимает единственный аргумент -
         *                                путь к файлу QRFile.
         *                                
         *                                
         *  >  public void PDFQRCodeRecognition(string InputFilePath)
         *                              - Функция распознавания QR-кода в PDF-файле. Не будет иметь параметров, т.к. путь ко входному
         *                                файлу уже задан.
         *                                
         *                               
         *  > private string QRCodeDecode(iTextSharp.text.Image QR)
         *                              - Функция расшифровки (считывания) QR-кода. Принимает один параметр - преобразованный
         *                                в Image считанный QR-код (возможно лучше преобразовать в BarcodeQRCode).
         * 
         * *****************************************************************************************************************/


        // Директория таблицы Excel, в которой хранятся расшифрованные QR-коды;
        private string DecocedQRCodesSpreadsheet = @"c:\QRPDF Test Directory\QRTest_DecodedQRCodes.xlsx";

        // Директория входного(ых) файла(ов)
        private string InputFile = @"c:\QRPDF Test Directory\QRTest_EmptyBlank.pdf";

        // Директория файла с информацией для QR-кода
        private string QRInfoFile = @"c:\QRPDF Test Directory\QRTest_QRInfo.txt";

        // Директория файла конфигурации программы
        private string ConfigurationsFile = @"c:\QRPDF Test Directory\QRTest_Configurations.txt";



        // Аксессор директории таблицы Excel с расшифрованными QR-кодами
        // Позволяет получить значение извне (Readonly)
        public string DecodedQRSpreadsheetPath
        {
            get { return DecocedQRCodesSpreadsheet; }
        }


        // Аксессор директории входного(ых) файла(ов)
        // Позволяет получить и установить новое значение извне (Read and Write)
        public string InputFilePath
        {
            get { return InputFile;  }
            set { InputFile = value; }
        }


        // Аксессор директории файла с информацией для QR-кода
        // Позволяет получить значение извне (Readonly)
        public string QRInfoFilePath
        {
            get { return QRInfoFile; }
        }


        // Аксессор директории файла конфигурации программы
        // Позволяет получить значение извне (Readonly)
        public string ConfigurationsFilePath
        {
            get { return ConfigurationsFile; }
        }


        // Типы документов
        public string[] DocumentsType = { "IN", "XZ", "MZ", "MM" };

        
        Random rand = new Random();


        // Приложение Excel
        public Excel.Application excel;
        public Excel.Worksheet sheet;
        public bool ExcelIsAlreadyRunning = false;







        /* ******************************************************************************************************************
         *                                                     ФУНКЦИИ
         * ******************************************************************************************************************/


        /* Конструктор по умолчанию. Входные данные вводятся с консоли:
         * 
         * Флаг -input  - путь к входному PDF-файлу;
         * Флаг -output - путь к итоговому PDF-файлу;
         * Флаг -qrfile - путь к файлу с информацией для генерации QR-кода;
         * Флаг -decode - декодируем QR-код входного файла;
         * Флаг -stamp  - установить новый QR-код;
         * Флаг -config - режим конфигурации (ручной ввод);
         * Флаг -help   - флаг для вывода напоминания о всех флагах.
         * 
         * ***********************************************************/
        public CQRPdf()
        {
            Console.Write("Input file path   -> ");
            string inp_InputFilePath = Console.ReadLine();

            Console.Write("Output file path  -> ");
            string inp_OutputFilePath = Console.ReadLine();

            Console.Write("QR text file path -> ");
            string inp_QRTextFilePath = Console.ReadLine();


            inp_InputFilePath = Regex.Replace(inp_InputFilePath, @"\t|\n|\r", "");
            inp_OutputFilePath = Regex.Replace(inp_OutputFilePath, @"\t|\n|\r", "");
            inp_QRTextFilePath = Regex.Replace(inp_QRTextFilePath, @"\t|\n|\r", "");



            this.InputFilePath = inp_InputFilePath;
            // this.OutputFilePath = inp_OutputFilePath;
            // this.QRInfoFilePath = inp_QRTextFilePath;
        }






        /* Конструктор с одним параметром - флагом. В зависимости от флага
         * 
         * В случае его вызова будет считан только QR-код входного файла.
         * 
         ********************************************************************/
        public CQRPdf(char flag)
        {

        }






        /*                    Конструктор с тремя параметрами:
         *              
         * - Директория входного PDF-файла;
         * - Директория выходного PDF-файла;
         * - Директория файла с информацией для генерации QR-кода;
         * 
         * (!) По возможности заменить директории на непосредственные файлы.
         * 
         * *******************************************************************/
        private CQRPdf(string pathToInputPDFDoc,
                       string pathToOutputPDFDoc,
                       string pathToQRTextFile)
        {
            InputFilePath = Regex.Replace(pathToInputPDFDoc, @"\t|\n|\r", "");
            // OutputFilePath = Regex.Replace(pathToOutputPDFDoc, @"\t|\n|\r", "");
            // QRInfoFilePath = Regex.Replace(pathToQRTextFile,   @"\t|\n|\r", "");
        }




        // Установка QR-кода в PDF-файл
        public void PDFStampQRCode(iTextSharp.text.Image QR,      // Сформированный QR-код (преобразованный из BarcodeQRCode в Image;
                                   int x_offset = 0,              // Отступ по оси х (по умолчанию 0   - левое верхнее положение);
                                   int y_offset = 750)            // Отступ по оси y (по умолчанию 750 - левое верхнее положение).
        {
            Document document = new Document();

            PdfWriter.GetInstance(document, new FileStream(InputFilePath, FileMode.OpenOrCreate, FileAccess.Write));
            document.Open();
            QR.SetAbsolutePosition(x_offset, y_offset);
            document.Add(QR);
            document.Close();
        }



        // Генерация нового QR-кода по считанной из файла информации
        public iTextSharp.text.Image QRGenerate(string QRFileContent)
        {
            BarcodeQRCode QRCode = new BarcodeQRCode(QRFileContent, 100, 100, null);

            Console.WriteLine(">> " + QRCode.GetHashCode());

            iTextSharp.text.Image QRCodeImage = QRCode.GetImage();

            return QRCodeImage;
        }



        // Добавить новую запись в базу
        public int Database_Add(string UniqueID, string Content)
        {
            Excel_AddNew(UniqueID, Content);
            return 0;
        }



        // Поиск документа в общей таблице обработанных документов 
        public void Database_FindDocumentByID(int UniqueID)
        {
            Excel_FindByID(UniqueID);
        }




        // Распознавание QR-кода в PDF-файле
        public void PDFQRCodeRecognition(string PDF)
        {
            //
        }




        // Расшифровка QR-кода
        private int QRCodeDecode(iTextSharp.text.Image QR)
        {
            /*
             *  PdfReader pdf = new PdfReader(filename);  
                PdfDictionary pg = pdf.GetPageN(pageNum);  
                PdfDictionary res = (PdfDictionary)PdfReader.GetPdfObject(pg.Get(PdfName.RESOURCES));  
                PdfDictionary xobj = (PdfDictionary)PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT));  
                
                if (xobj == null) { return; }
                
                foreach (PdfName name in xobj.Keys)  
                {  
                    PdfObject obj = xobj.Get(name);  
                    if (!obj.IsIndirect()) { continue; }  
                   
                    PdfDictionary tg = (PdfDictionary)PdfReader.GetPdfObject(obj);  
                    PdfName type = (PdfName)PdfReader.GetPdfObject(tg.Get(PdfName.SUBTYPE));  
                    
                    if (!type.Equals(PdfName.IMAGE)) { continue; }  
                   
                    int XrefIndex = Convert.ToInt32(((PRIndirectReference)obj).Number.ToString(System.Globalization.CultureInfo.InvariantCulture));  
                    PdfObject pdfObj = pdf.GetPdfObject(XrefIndex);  
                    PdfStream pdfStrem = (PdfStream)pdfObj;  
                    byte[] bytes = PdfReader.GetStreamBytesRaw((PRStream)pdfStrem);  
                    
                    if (bytes == null) { continue }; 
                   
                    memStream.Position = 0;  
                    System.Drawing.Image img = System.Drawing.Image.FromStream(memStream);
                    
                }
  
                string path = Path.Combine(String.Format(@"result-{0}.jpg", pageNum));  
                System.Drawing.Imaging.EncoderParameters parms = new System.Drawing.Imaging.EncoderParameters(1);  
                parms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Compression, 0);  
                var jpegEncoder = ImageCodecInfo.GetImageEncoders().ToList().Find(x => x.FormatID == ImageFormat.Jpeg.Guid);  
                img.Save(path, jpegEncoder, parms);
             */
            return 0;
        }



        public string GenerateUniqueID()
        {
            return Convert.ToString(DocumentsType[rand.Next(0, 4)] + rand.Next(1000000, 9999999));
        }
        


        // Запуск Excel'a. Инициализация приложения, книг, листов и т.д.
        public int Excel_RunApplication()
        {
            if (ExcelIsAlreadyRunning == false)
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                excel.SheetsInNewWorkbook = 1;
                excel.DisplayAlerts = false;
                Excel.Workbook workBook = excel.Workbooks.Open(DecodedQRSpreadsheetPath);
                sheet = (Excel.Worksheet)excel.Worksheets.get_Item(1);

                Excel_SetConfiguration();
                ExcelIsAlreadyRunning = true;
            } else
            {
                return -1;
            }

            return 0;
        }



        // Сохранение таблицы
        public int Excel_Save()
        {
            if (ExcelIsAlreadyRunning == false)
                Excel_RunApplication();

            excel.Application.ActiveWorkbook.Save();
            return 0;
        }



        // Поиск элемента по ID
        public string Excel_FindByID(int UniqueID)
        {
            if (ExcelIsAlreadyRunning == false)
                Excel_RunApplication();

            return sheet.Cells[UniqueID, 2];
        }



        // Поиск первой пустой ячейки (для добавления новых элементов)
        public int Excel_FindFirstEmpty()
        {
            if (ExcelIsAlreadyRunning == false)
                Excel_RunApplication();

            int lastRow, rowsCount = sheet.Rows.Count;

            Excel.Range xlRange = (Excel.Range)sheet.Cells[rowsCount, 1];
            lastRow = xlRange.get_End(Excel.XlDirection.xlUp).Row;


            // Автоматическая настройка ширины и высоты
            sheet.Cells[rowsCount, 1].EntireColumn.AutoFit();
            sheet.Cells[rowsCount, 1].EntireRow.AutoFit();
            sheet.Cells[rowsCount, 2].EntireColumn.AutoFit();
            sheet.Cells[rowsCount, 2].EntireRow.AutoFit();


            return ++lastRow;
        }




        // Добавление нового элемента
        public int Excel_AddNew(string UniqueID, string Content)
        {
            if (ExcelIsAlreadyRunning == false)
                Excel_RunApplication();

            int index = Excel_FindFirstEmpty();

            sheet.Cells[index, 1] = UniqueID;
            sheet.Cells[index, 2] = Content;

            return 0;
        }



        // Конфигурация Excel таблицы
        private void Excel_SetConfiguration()
        {
            // Название листа
            sheet.Name = "QR Test List 1";            // Название листа;


            // Установка свойств верхней строки
            sheet.Cells[1, 1] = "ID";                 // Заголовок первого столбца;
            sheet.Cells[1, 2] = "Code";               // Заголовок второго столбца;

            // Установка цвета фона для верхней строки
            sheet.Cells[1, 1].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0x91, 0x91, 0x91));
            sheet.Cells[1, 2].Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0x91, 0x91, 0x91));

            // Установка шрифта и выравнивания
            // Выравнивание по центру
            sheet.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            sheet.Cells.Font.Name = "Century Gothic"; // Установка шрифта;
            sheet.Cells.Font.Size = 10;               // Установка размера шрифта.

        }
    }
}

