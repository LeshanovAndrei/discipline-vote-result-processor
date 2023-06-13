using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DVRP1
{
    class Selection
    {
        public Selection(Discipline discipline)
        {
            this.discipline = discipline;
            Students = new List<Student>();
            status = 0;

        }
        public void AddStudent(Student st)
        {
            Students.Add(st);
        }

        public void StatusCheck()
        {
                status = 0;
            if (students.Count >= discipline.ForConditChoose)
                status = 1;
            if (students.Count >= discipline.ForChoose)
                status = 2;
        }
        private Discipline discipline;
        private List<Student> students;
        private uint status;
        private bool magistr;
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

        internal List<Student> Students { get => students; set => students = value; }
        public bool Magistr { get => magistr; set => magistr = value; }
    }
}
