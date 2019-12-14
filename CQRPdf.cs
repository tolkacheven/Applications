using System;
using System.IO;
using System.Drawing.Printing;
using System.Text.RegularExpressions;
using PQScan.PDFToImage;
using PQScan.BarcodeScanner;
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
         *  >  Поточно обрабатывать переданные файлы;
         *  >  Сохранять полученные данные в Excel-таблице;
         *  >  
         *  
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
         *     >  Configurations[0] - Stamp         Mode (-stamp)     - режим установки QR-кода;
         *     >  Configurations[1] - Stamp  All    Mode (-stampall)  - режим поточной установки QR-кода;
         *     >  Configurations[2] - Decode        Mode (-decode)    - режим расшифровки QR-кода;
         *     >  Configurations[3] - Decode All    Mode (-decodeall) - режим поточной расшифровки QR-кода;
         *     >  Configurations[4] - Configuration Mode (-config)    - режим конфигурации программы.
         *
         * 
         * 
         * ******************************************************************************************************************
         *                                             ПРОГРАММНЫЙ ИНТЕРФЕЙС
         * ******************************************************************************************************************
         * 
         * #######################################
         * ###       Конструкторы класса       ###
         * #######################################
         * 
         *     >  public CQRPdf()          - Конструктор по умолчанию. Создает экземпляр класса с предустановленными значениями;
         *  
         *                                
         *                                
         *     >  private CQRPdf(string,   - Конструктор, вызываемый только режимом конфигурации, все три параметра обозначают три директории:
         *                       string,     InputFilePath, DecodedQRCodesSpreadsheet, QRInfoFilePath.
         *                       string)
         *                       
         *         
         *         
         *         
         * ########################################
         * ###   Создание и установка QR-кода   ###
         * ########################################
         *                    
         *     >  public void PDFStampQRCode(iTextSharp.text.Image QR,
         *                                   int x_offset = 0,   
         *                                   int y_offset = 750)
         *                                
         *                                   - Функция установки QR-кода. Принимает три параметра:
         *                                     > iTextSharp.text.Image QR - преобразованный QR-код (из типа BarcodeQRCode в Image);
         *                                     > x_offset - опциональный параметр (по умолчанию   0) - отступ QR-кода по оси x;
         *                                     > y_offset - опциональный параметр (по умолчанию 750) - отступ QR-кода по оси y.
         *                                
         *                                
         *     >  public iTextSharp.text.Image QRGenerate(string QRFileContent)
         *                                   - Функция, генерирующая QR-код по содержимому файла. Принимает единственный аргумент -
         *                                     путь к файлу QRFile.
         *                                     
         *                                     
         *                                     
         *                                     
         * ########################################
         * ###       Распознавание QR-кода      ###       
         * ########################################
         * 
         *       >  public void PDFQRCodeRecognition(string InputFilePath)
         *                              - Функция распознавания QR-кода в PDF-файле. Не будет иметь параметров, т.к. путь ко входному
         *                                файлу уже задан.
         *                                
         *                               
         *       >  private string QRCodeDecode(iTextSharp.text.Image QR)
         *                                   - Функция расшифровки (считывания) QR-кода. Принимает один параметр - преобразованный
         *                                     в Image считанный QR-код (возможно лучше преобразовать в BarcodeQRCode).
         *                                     
         *                                     
         *                                     
         *                                     
         * ########################################
         * ###     Работа с таблицами Excel     ###
         * ########################################
         *       
         *       > public int Excel_RunApplication()     - Запуск Excel'a. Инициализация приложения, книг, листов и т.д;
         *      
         *       > private void Excel_SetConfiguration() - Конфигурация Excel таблицы. Установка шрифтов, наименований
         *                                                 колонок и т.д;
         *         
         *       > private int Excel_Save()              - Сохранение таблицы;
         *       
         *       > public string Excel_FindByID (string UniqueID)
         *                                               - Поиск в таблице по уникальному ID;
         *                                               
         *       > private int Excel_FindFirstEmpty()    - Поиск первой пустой строки (нужен для добавления новых элементов);
         *       
         *       > public int Excel_AddNew(string UniqueID, string Content)
         *                                               - Добавить новый элемент в таблицу Excel.
         *                  
         *                  
         *                  
         *                  
         *                  
         * ########################################
         * ###        Работа с принтером        ###
         * ########################################
         * 
         * 
         * *****************************************************************************************************************/







        /******************************************************************************************************************
         *                                                    ДИРЕКТОРИИ
         ******************************************************************************************************************/

        // Корневая директория
        static private string rootDirectory      = @"c:\QRPDF Test Directory";

        // Директория таблицы Excel, в которой хранятся расшифрованные QR-коды;
        private string DecocedQRCodesSpreadsheet = rootDirectory + @"\QRTest_DecodedQRCodes.xlsx";

        // Директория входного(ых) файла(ов)
        private string InputFile                 = rootDirectory + @"\QRTest_blank.pdf";

        // Директория файла с информацией для QR-кода
        private string QRInfoFile                = rootDirectory + @"\QRTest_QRInfo.txt";
                
        // Директория файла конфигурации программы
        private string ConfigurationsFile        = rootDirectory + @"\QR Configuration Presets\QRTest_Presets_Default.ini";

        // Директория с файлами для поточной обработки
        private string MassProcessingDirectory   = rootDirectory + @"\QR Mass Processing Mode Test\";







        /******************************************************************************************************************
         *                                                АКСЕССОРЫ ДИРЕКТОРИЙ
         ******************************************************************************************************************/


        // Аксессор корневой директории
        public string RootDirectoryPath
        {
            get { return rootDirectory;  }
            set { rootDirectory = value; }
        }


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



        // Аксессор директории с файлами, для многопоточной обработки
        // Позволяет получить и установить новое значение извне (Read and Write)
        public string MassProcessingDirectoryPath
        {
            get { return MassProcessingDirectory;  }
            set { MassProcessingDirectory = value; }
        }







        /******************************************************************************************************************
         *                                                  КОНФИГУРАЦИЯ
         ******************************************************************************************************************/


        // Конфигурация программы. Флаги
        protected bool[] Configurations = { 
                                            /* Stamp       -> */ false,
                                            /* Stamp  All  -> */ false,
                                            /* Decode      -> */ false,
                                            /* Decode All  -> */ false,
                                            /* Config      -> */ false
                                          };


        
        // Типы документов
        public string[] DocumentsType = { "IN", "XZ", "MZ", "MM" };


        





        /******************************************************************************************************************
         *                                                  РАБОТА С EXCEL
         ******************************************************************************************************************/

        public Excel.Application excel;                 // Объект для работы с функциями приложения; 

        public Excel.Worksheet   sheet;                 // Объект для работы с функциями листа;

        public bool ExcelIsAlreadyRunning = false;      // Флаг, отвечающий за текущее состояние экселя.








        /******************************************************************************************************************
         *                                               РАБОТА С ПРИНТЕРОМ
         ******************************************************************************************************************/








        /******************************************************************************************************************
         *                                                    ОСТАЛЬНОЕ
         ******************************************************************************************************************/

        // Переменная для генерации случайных чисел
        Random rand = new Random();








        /* ******************************************************************************************************************
         *                                                     ФУНКЦИИ
         * ****************************************************************************************************************** */


        /*                   Конструктор по умолчанию
         * 
         * Создает экземпляр класса с предустановленными значениями;
         * 
         * ***********************************************************/
        public CQRPdf()
        {
            int currentRow = 0;
            string currentLine;
            string path = ConfigurationsFilePath;
            
            using (StreamReader reader = new StreamReader(path))
            {
                while (reader.Peek() >= 0)
                {
                    currentLine = reader.ReadLine();
                    if (currentLine == "0") Configurations[currentRow++] = false; else Configurations[currentRow++] = true;
                }
            }
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
        public CQRPdf(string pathToInputPDFDoc,
                      string pathToOutputPDFDoc,
                      string pathToQRTextFile)
        {
            InputFilePath = Regex.Replace(pathToInputPDFDoc, @"\t|\n|\r", "");
            // OutputFilePath = Regex.Replace(pathToOutputPDFDoc, @"\t|\n|\r", "");
            // QRInfoFilePath = Regex.Replace(pathToQRTextFile,   @"\t|\n|\r", "");
        }




        // Установка QR-кода в PDF-файл
        public void PDFStampQRCode(iTextSharp.text.Image QR,      // Изображение (преобразованное из BarcodeQRCode в Image);
                                            int xOffset = 0,      // Отступ по оси х. За начало отсчета считается левый нижний угол;
                                            int yOffset = 740,    // Отступ по оси y. За начало отсчета считается левый нижний угол;
                                            string FilePath = "c:\\QRPDF Test Directory\\QRTest_blank.pdf")    
        {

            string CopyFileName = FilePath.Substring(0, FilePath.Length - 4) + "_QR.pdf";

            var reader     = new PdfReader(FilePath);
            var fileStream = new FileStream(CopyFileName, FileMode.Create, FileAccess.Write);
            var document   = new Document(reader.GetPageSizeWithRotation(1));
            var writer     = PdfWriter.GetInstance(document, fileStream);

            document.Open();

            for (var i = 1; i <= reader.NumberOfPages; i++)
            {
                document.NewPage();
                                
                var importedPage = writer.GetImportedPage(reader, i);
                var contentByte  = writer.DirectContent;

                contentByte.BeginText();
                QR.SetAbsolutePosition(xOffset, yOffset);
                contentByte.AddImage(QR);
                contentByte.EndText();
                contentByte.AddTemplate(importedPage, 0, 0);
            }

            document.Close();
            writer.Close();

        }




        // Установка QR-кодов на все PDF-файлы в директории
        public void PDFStampQRCode_Mass(iTextSharp.text.Image QR,      // Изображение (преобразованное из BarcodeQRCode в Image);
                                                 int xOffset = 0,      // Отступ по оси х. За начало отсчета считается левый нижний угол;
                                                 int yOffset = 755)    // Отступ по оси y. За начало отсчета считается левый нижний угол;
        {
            string FilePath = MassProcessingDirectoryPath;
            string tmpString;
            int tmpIndex;

            DirectoryInfo directory = new DirectoryInfo(FilePath);


            foreach (var file in directory.GetFiles("*.pdf"))
            {
                tmpIndex  = 0;
                tmpString = file.FullName;

                foreach (char letter in tmpString)
                {
                    if (letter == '\\')
                    {
                        tmpString = tmpString.Insert(tmpIndex, "\\");
                        tmpIndex++;
                    }

                    tmpIndex++;

                }
                this.PDFStampQRCode(QR, xOffset, yOffset, tmpString);  
            }
        }






        // Генерация нового QR-кода по считанной из файла информации
        public iTextSharp.text.Image QRGenerate(string QRFileContent)
        {
            BarcodeQRCode QRCode = new BarcodeQRCode(QRFileContent, 90, 90, null);
            iTextSharp.text.Image QRCodeImage = QRCode.GetImage();

            //QRCodeImage.SetDpi(1300, 1300);
            
            return QRCodeImage;
        }
               



        // Распознавание QR-кода в PDF-файле
        public void PDFQRCodeRecognition(string PDF)
        {
            Bitmap bmp = new Bitmap("c:\\QRPDF Test Directory\\Cutted QR\\Test11.png");
            BarcodeResult barcode = BarCodeScanner.ScanSingle(bmp);
            Console.WriteLine("barcode data:{0}.", barcode.Data);

            
            BarcodeResult[] results = BarCodeScanner.Scan("c:\\QRPDF Test Directory\\Cutted QR\\Test11.png", BarCodeType.QRCode);

            foreach (BarcodeResult result in results)
            {
                Console.WriteLine(result.BarType.ToString() + "-" + result.Data);
            }

            Console.ReadKey();
        }




        // Расшифровка QR-кода
        private int QRCodeDecode(iTextSharp.text.Image QR)
        {
            PdfReader pdf      = new PdfReader(InputFilePath);  
            PdfDictionary pg   = pdf.GetPageN(1);  
            PdfDictionary res  = (PdfDictionary)PdfReader.GetPdfObject(pg.Get(PdfName.RESOURCES));  
            PdfDictionary xobj = (PdfDictionary)PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT));  
            
            if (xobj == null) { return 0; }
            
            foreach (PdfName name in xobj.Keys)  
            {  
                PdfObject obj = xobj.Get(name);  
                if (!obj.IsIndirect()) { continue; }  
               
                PdfDictionary tg = (PdfDictionary)PdfReader.GetPdfObject(obj);  
                PdfName type = (PdfName) PdfReader.GetPdfObject(tg.Get(PdfName.SUBTYPE));  
                
                if (!type.Equals(PdfName.IMAGE)) { continue; }  
               
                int XrefIndex = Convert.ToInt32(((PRIndirectReference)obj).Number.ToString(System.Globalization.CultureInfo.InvariantCulture));  
                PdfObject pdfObj = pdf.GetPdfObject(XrefIndex);  
                PdfStream pdfStrem = (PdfStream)pdfObj;  
                byte[] bytes = PdfReader.GetStreamBytesRaw((PRStream)pdfStrem);  
                
                if (bytes == null) { continue; }; 
               
                //memStream.Position = 0;  
                //System.Drawing.Image img = System.Drawing.Image.FromStream(memStream);
                
            }
  
            string path = Path.Combine(String.Format(@"result-{0}.jpg", 1));  
            System.Drawing.Imaging.EncoderParameters parms = new System.Drawing.Imaging.EncoderParameters(1);  
            parms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Compression, 0);  
            var jpegEncoder = ImageCodecInfo.GetImageEncoders().ToList().Find(x => x.FormatID == ImageFormat.Jpeg.Guid);  
            //img.Save(path, jpegEncoder, parms);

            return 0;
        }




        // Генерирование уникального ID
        public string GenerateUniqueID()
        {
            int randomNumber = rand.Next(1, 9999999);
            string resultID  = DocumentsType[rand.Next(0, 4)];

            for (int i = 0; i < 7 - randomNumber.ToString().Length; i++) resultID += "0";

            resultID += randomNumber.ToString();

            return resultID;
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
            }
            else
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
        public string Excel_FindByID (string UniqueID)
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

            Excel.Range xlRange = (Excel.Range) sheet.Cells[rowsCount, 1];
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
            sheet.Cells.Font.Name = "Tahoma";         // Установка шрифта;
            sheet.Cells.Font.Size = 10;               // Установка размера шрифта.

        }        
    }
}



