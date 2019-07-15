using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace excel2
{
    class Program
    {
        static void Main(string[] args)
        {
           // DateTime start = DateTime.Now;

            using (var stream = File.Open("C:/Users/LENOVO/Desktop/excele/101103.xls", FileMode.Open, FileAccess.Read))
            {

                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    using (var SW = new StreamWriter("C:/Users/LENOVO/Desktop/excele/Nowy.csv"))
                    {
                            while (reader.Read())
                            {

                                int size = (int)reader.FieldCount;
                                for (int i = 0; i < size; i++)
                                {
                                    string Cals = Convert.ToString(reader.GetValue(i));

                                    if (Cals.Contains("\n"))
                                    {
                                        Cals = Cals.Replace("\n", " ");
                                    }
                                    Console.Write(Cals + "  ");
                                    SW.Write(Cals + ";");
                                }
                                SW.WriteLine();
                                 Console.Write("\n");
                            }
                            // } while (reader.NextResult());

                        reader.Close();

                       // DateTime stop = DateTime.Now;
                       // Console.WriteLine("time" + (stop - start).ToString());
                    }

                    Console.ReadKey();
                }
            }
        }
    }
}
