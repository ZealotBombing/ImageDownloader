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
            string[] paths = new string[]
            {
                @"C:\Users\accio\OneDrive\Escritorio\Respaldo\Unne\Multimedia\propMultimedia_0181bD1LcCYLacEdL1FKaQF8HRQSFff4U0EQAKMcUC6g3UDPeKQ3M9.xlsx",
                @"C:\Users\accio\OneDrive\Escritorio\Respaldo\Unne\Multimedia\propMultimedia_0181bX09KRW4VZ4RG7MKUJ1Af7ebORZUVNC22JIDPIZFF0HI5OUXfC.xlsx",
                @"C:\Users\accio\OneDrive\Escritorio\Respaldo\Unne\Multimedia\propMultimedia_0181LWEQJVNgCPbCJSUeBe6ILfcLWTE7LcXa7HQ3gOAaae44FXMIE7.xlsx",
                @"C:\Users\accio\OneDrive\Escritorio\Respaldo\Unne\Multimedia\propMultimedia_0181GeDMXTQKFUVdeZVFRJUObBBPJHAMIYHOfA5LZZSCVc0eDCg3Og.xlsx",
      
            };

            Dictionary<string, Dictionary<string, string>> data = ReadExcelFile(paths[3]);

            GuardarImagen(data);
        }

        public static void GuardarImagen(Dictionary<string, Dictionary<string, string>> data)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {

                    foreach (KeyValuePair<string, Dictionary<string, string>> docSet in data)
                    {
                        string propPath = $@"C:\Users\accio\OneDrive\Escritorio\Respaldo\Test\Img\{docSet.Key.Replace(".","")}";//Panal 2
                        //string propPath = $@"C:\Users\cfuen\OneDrive\Escritorio\Imagenes pásivas\Img\{docSet.Key.Replace(".", "")}";//Panal 2
                        Console.WriteLine(docSet.Key);

                        if (!Directory.Exists(propPath))
                        {
                            Directory.CreateDirectory(propPath);
                        }

                        if (Directory.GetFiles(propPath).Length >= docSet.Value.Count)
                        {
                            continue;
                        }

                        int count = 0;

                        foreach (KeyValuePair<string, string> doc in docSet.Value)
                        {
                            if (doc.Value == "Imagen")
                            {
                                string url = doc.Key;
                                HttpResponseMessage response = client.GetAsync(url).GetAwaiter().GetResult();

                                Console.WriteLine(doc.Key);

                                if (!response.IsSuccessStatusCode)
                                {
                                    Console.WriteLine(response.StatusCode);
                                    Console.BackgroundColor = ConsoleColor.Red;
                                    continue;
                                };

                                byte[] imageBytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();

                                File.WriteAllBytes($@"{propPath}\{count}.jpeg", imageBytes);

                                count++;
                            }
                        }
                    }

                    Beep();

                    Console.WriteLine("Ok");
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }
            finally
            {
                Beep();
            }
        }

        static Dictionary<string, Dictionary<string, string>> ReadExcelFile(string filePath)
        {

            //string filePath = @"C:\Users\accio\OneDrive\Escritorio\Respaldo\Unne\Multimedia\propMultimedia_0181bD1LcCYLacEdL1FKaQF8HRQSFff4U0EQAKMcUC6g3UDPeKQ3M9.xlsx";


            //ExcelPackage.LicenseContext = LicenseContext.Commercial;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            FileInfo fileInfo = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                Dictionary<string, Dictionary<string, string>> data = new Dictionary<string, Dictionary<string, string>>();

                string idProp = "";

                for (int row = 91461; row <= rowCount; row++)
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
                    }
                }

                return data;
            }
        }

        public static void Beep()
        {
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();
        }
    }
}