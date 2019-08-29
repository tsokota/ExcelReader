using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace DataReader.SummerPractice
{

    public class DataReader
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        public List<MarkRecord> ReadMarks(string path)
        {
            // подключаемся к файлу
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(path);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            var lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            List<MarkRecord> markRecords = new List<MarkRecord>();

            Console.WriteLine("Чтение данных...");
            // считываем массив записей 
            for (int index = 2; index <= lastRow; index++)
            {
                try
                {
                    System.Array MyValues = (System.Array)MySheet.get_Range("A" +
                       index.ToString(), "F" + index.ToString()).Cells.Value;


                    markRecords.Add(new MarkRecord
                    {
                        StudentsId = Convert.ToInt32(MyValues.GetValue(1, 1)),
                        Sername = MyValues.GetValue(1, 2).ToString(),
                        Name = MyValues.GetValue(1, 3).ToString(),
                        Group = MyValues.GetValue(1, 4).ToString(),
                        SubjectName = MyValues.GetValue(1, 5).ToString(),
                        SetMarkValue = MyValues.GetValue(1, 6).ToString()
                    });
                }
                catch(Exception E)
                {
                    break;
                }
                if (markRecords.Last().StudentsId > 0) { }
                else break;
            }

            return markRecords;
        }

        public decimal GetPercentBetterStudents(string pathF, string pathT)
        {
            List<MarkRecord> markRecords = ReadMarks(pathF);

            int counter = 0;

            // солучаем массив оценок, сгруппированы по студентам
            List<IGrouping<int, MarkRecord>> students = markRecords.GroupBy(x => x.StudentsId).ToList();

            Console.WriteLine("Обработка...");
            // считаем кол-во студентов, которые имеют только 4-5
            students.ForEach(x => {
                {
                    bool tmp = false;
                    foreach (var y in x)
                    {
                        if (y.MarkValue < 75)
                        {
                            tmp = true;
                        }
                    }
                    if (!tmp)
                    {
                        counter++;
                    }

                }
            });

            //считаем % студентов с оценками 4-5
            decimal result = (decimal)((decimal)100 / students.Count) * counter;

            // записуем результат в файл
            WriteResult(result, pathT);

            return result;
        }

        public void WriteResult(decimal res, string pathT)
        {
            // подключаемся к файлу
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(pathT);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            var lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            Console.WriteLine("Сохранение результатов в файл...");
            lastRow += 1;
            MySheet.Cells[lastRow, 1] = lastRow - 1 ;
            MySheet.Cells[lastRow, 2] = DateTime.Now.ToString();
            MySheet.Cells[lastRow, 3] = res;
           // EmpList.Add(emp);
            MyBook.Save();
        }
    }
}
