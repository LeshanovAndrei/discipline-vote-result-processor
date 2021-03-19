using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DVRP1
{
    class Selection
    {
        public void AddStudent(Student st)
        {
            students.Add(st);
            numbers++;
            //Какая-нибудь проверка статуса
        }
        private Discipline discipline;
        private List<Student> students;
        private uint numbers;
        private uint status;
        public Discipline Discipline
        {
            get
            {
                return discipline;
            }
            set
            {
                discipline = value;
            }
        }
        
        public uint Numbers
        {
            get
            {
                return numbers;
            }
            set
            {
                numbers = value;
            }
        }
        
        public uint Status
        {
            get
            {
                return status;
            }
            set
            {
                status = value;
            }
        }
    }
}
