using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataReader.SummerPractice
{
    public class MarkRecord
    {
        public int StudentsId { get; set; }

        public string Name{ get; set; }

        public string Sername { get; set; }

        public string Group { get; set; }

        public int MarkValue { get; set; }

        public string SubjectName { get; set; }

        public string SetMarkValue {
            set
            {
                int tmp = 0;
                Int32.TryParse(value, out tmp);
                this.MarkValue = tmp;
            }
        }
    }
}
