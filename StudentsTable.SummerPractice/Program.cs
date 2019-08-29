using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using nmsp = DataReader.SummerPractice;

namespace StudentsTable.SummerPractice
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            // расположение .xls файла с данными
            string pathFrom = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "STUDENT_MARKS.xls");
          
            // расположение .xls файла для результата
            string pathTo = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results.xlsx");

            Console.WriteLine("Старт программы..");
            nmsp.DataReader dataReader = new nmsp.DataReader();
            try
            {
                Console.WriteLine("Результат: " + dataReader.GetPercentBetterStudents(pathFrom, pathTo).ToString("0.00")+"%");
                Console.WriteLine("Успешно!"); 
            }
            catch(Exception ex)
            {
                Console.WriteLine("Ошибка! Проверьте все ли впрорядке с файлами .xls"+ ex.Message);
            }
           
            Console.Read();
        }
    }
}
