using System;
using System.IO;
using System.Security.AccessControl;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QRPDF
{
    class Program
    {
        static void Main(string[] args)
        {            
            CQRPdf TestModule;
            
            
            if (args.Count() == 0)                      // Если нет дополнительных аргументов - создаем экземпляр класса CQRPdf,                     
            {                                           // используя конструктор по умолчанию;

                // 12345678912345678912345678912345678912 - 38 символов;
                // ABCDEFGHKLMNOPR123456789123 - 27 символов
                // QRPDFR12345678912345678


                TestModule = new CQRPdf();

                //TestModule.PDFStampQRCode(TestModule.QRGenerate(File.ReadAllText(TestModule.QRInfoFilePath, Encoding.UTF8)));

                TestModule.PDFQRCodeRecognition("");

                //TestModule.PDFStampQRCode_Mass(TestModule.QRGenerate(File.ReadAllText(TestModule.QRInfoFilePath, Encoding.UTF8)));                

                //for (int i = 0; i < 20; i++) TestModule.Excel_AddNew(TestModule.GenerateUniqueID(), "Test #" + i + " // QR Content");

                //TestModule.Excel_Save();
                               


            } else if (args.Count() % 2 != 0) {         // Если количество аргументов нечетно - считаем, что был передан флаг -help,
                                                        // т.к. все остальные требуют указания директории (а значит, чётны).
                                                        // Пока считаем, что в случае, если указан не флаг -help - мы прерываем программу.
                if (args[0] == "-help")
                {
                    Console.WriteLine("### HELP\n"                       +
                                      "-out    - итоговый файл;\n"       +
                                      "-inp    - входной файл;\n"        +
                                      "-decode - расшифровать QR-код;\n" +
                                      "-stamp  - установить QR-код;\n"   +
                                      "-qrfile - файл с информацией для QR-кода\n");

                    TestModule = new CQRPdf();

                } else {

                    return;

                }

            } else {

                /* 
                 * Итератор. Предназначен для отображения позиции текущего флага в строке. Такой
                 * подход позволит располагать флаги в произвольном порядке, а также сразу получать
                 * информацию о необходимой директории (по обращению к следующему элементу.
                 */
                int filePaths = 0;


                // Переменные для временного хранения путей к файлам
                string inp_InputFilePath  = "",     // Директория входного файла;
                       inp_OutputFilePath = "",     // Директория выходного файла;
                       inp_QRTextFilePath = "";     // Директория файла с информацией для QR-кода.
               

                foreach (string arguments in args)
                {
                    filePaths++;
                    switch (arguments)
                    {
                        case "-out":    Console.WriteLine("Выходной файл: " + args[filePaths]); inp_InputFilePath  = args[filePaths]; break;
                        case "-inp":    Console.WriteLine("Входной файл: "  + args[filePaths]); inp_OutputFilePath = args[filePaths]; break;
                        case "-qrfile": Console.WriteLine("Выходной файл: " + args[filePaths]); inp_QRTextFilePath = args[filePaths]; break;
                        case "-help":   Console.WriteLine("-out - итоговый файл;\n-file - входной файл;" +
                                                          "\n-qrfile - файл с информацией для QR-кода\n"); break;
                    }
                }                
            }
        }     
    }
}
