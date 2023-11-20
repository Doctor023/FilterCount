using NPOI.HSSF.Model;
using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace FilterCount
{
    internal class Program
    {
        static List<string> _inputValues = new List<string>();

        static int _limit;
        private static void Main(string[] args)
        {
            Console.WriteLine("Hi, fill out the file Input.txt and save it, then enter value and I will give you only the result that matches the given value for the number of characters and save it in a folder Results");
            Console.WriteLine("Enter a character limit");

            while (true)
            {
                string? limit = Console.ReadLine();
                if (!int.TryParse(limit, out _limit))
                {
                    Console.WriteLine("I can get only numbers");
                    continue;
                }
                break;
            }

            string filePath = @"Input.txt";

            string[] lines = File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                if (line.Count() == _limit)
                {
                    _inputValues.Add(line);
                }
            }
            CreateFile();
        }
  
        static void CreateFile()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Лист1");
            int i = 0;
            foreach (var value in _inputValues)
            {
                IRow row = sheet.CreateRow(i);
                ICell cell0 = row.CreateCell(0);
                cell0.SetCellValue(value);
                Console.WriteLine(value);
                i++;
            }

            string filePath = "File1";
            for (int c = 1; true; c++)
            {
                if (File.Exists(@$"Results\File{c}.xlsx"))
                {
                    continue;
                }
                filePath = "File" + c;
                break;
                    
            }


            using (FileStream file = new FileStream($@"Results\{filePath}.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
                file.Close();
            }
        }
    }
}