using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, Dictionary<string, string>> data = ReadExcelFile();

            GuardarImagen(data);
        }


        public static void GuardarImagen(Dictionary<string, Dictionary<string, string>> data)
        {

            //string url = "https://demoazimg.prop360.cl/unne/img/propiedades/79772_clkSAxQjeh5hs6afOvtl20240624001706.jpeg";
            string path = @"C:\Users\accio\OneDrive\Escritorio\Respaldo\Test\Img";

            try
            {


                using (HttpClient client = new HttpClient())
                {

                    foreach(KeyValuePair<string,Dictionary<string, string>> docSet in data)
                    {
                        string propPath = $@"C:\Users\accio\OneDrive\Escritorio\Respaldo\Test\Img\{docSet.Key}";



                        if (!Directory.Exists(propPath))
                        {
                            Directory.CreateDirectory(propPath);

                        }

                        int count = 0;

                        foreach (KeyValuePair<string, string> doc in docSet.Value)
                        {
                            if (doc.Value == "Imagen")
                            {
                                string url = doc.Key;
                                HttpResponseMessage response = client.GetAsync(url).GetAwaiter().GetResult();

                                response.EnsureSuccessStatusCode();

                                byte[] imageBytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();

                                File.WriteAllBytes($@"{propPath}\{count}.jpeg", imageBytes);

                                count++;
                            }
                        }
                    }

                    Console.WriteLine("Ok");
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
        }

        public static void GuardarImagen()
        {

            string url = "https://demoazimg.prop360.cl/unne/img/propiedades/79772_clkSAxQjeh5hs6afOvtl20240624001706.jpeg";
            string path = @"C:\Users\accio\OneDrive\Escritorio\Respaldo\Test\Img";

            try
            {

                using (HttpClient client = new HttpClient())
                {
                    HttpResponseMessage response = client.GetAsync(url).GetAwaiter().GetResult();

                    response.EnsureSuccessStatusCode();

                    byte[] imageBytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();


                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);

                    }

                    File.WriteAllBytes($@"{path}\2.jpeg", imageBytes);

                    Console.WriteLine("Ok");

                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
        }

        //static void ReadExcelFile(string filePath)
        static Dictionary<string, Dictionary<string, string>> ReadExcelFile()
        {

            string filePath = @"C:\Users\accio\OneDrive\Escritorio\Respaldo\Unne\Multimedia\propMultimedia_0181bD1LcCYLacEdL1FKaQF8HRQSFff4U0EQAKMcUC6g3UDPeKQ3M9.xlsx";

            
            //ExcelPackage.LicenseContext = LicenseContext.Commercial;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            FileInfo fileInfo = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                // Get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Get the number of rows and columns in the worksheet
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                Dictionary<string, Dictionary<string, string>> data = new   Dictionary<string, Dictionary<string, string>>();

                string idProp = "";

                // Loop through the worksheet rows and columns
                for (int row = 7; row <= rowCount; row++)
                {

                    for (int col = 1; col <= colCount; col++)
                    {
                        if (col == 2)
                        {
                            idProp = worksheet.Cells[row, col].Text;
                            if (!data.ContainsKey(idProp))
                            {
                                data.Add(idProp, new Dictionary<string, string>());
                            }
                            
                        }
                        else if (col == 4)
                        {
                            if (!data[idProp].ContainsKey(worksheet.Cells[row, col].Text))
                            {
                                data[idProp].Add(worksheet.Cells[row, col].Text, worksheet.Cells[row, 3].Text);
                            }
                        }

                        //var cellValue = worksheet.Cells[row, col].Text;
                        //Console.Write($"{cellValue}\t");
                    }
                    //Console.WriteLine();
                }

                return data;
            }

        }
    }
}
