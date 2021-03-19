using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DVRP1
{
    class Student
    {
        public Student(string faculty, string email, string name, string course, string group, uint numbersOfDisc, List<string> codes)
        {
            this.faculty = faculty;
            this.email = email;
            this.name = name;
            this.course = course;
            this.group = group;
            this.numbersOfDisc = numbersOfDisc;
            this.codes = codes;
        }

        private string faculty;
        public string Faculty
        {
            get
            {
                return faculty;
            }
            set
            {
                faculty = value;
            }
        }

        private string email;

        public string Email
        {
            get
            {
                return email;
            }
            set
            {
                email = value;
            }
        }

        private string name;
        public string Name
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        private string course;
        public string Course
        {
            get
            {
                return course;
            }
            set
            {
                course = value;
            }
        }

        private string group;
        public string Group
        {
            get
            {
                return group;
            }
            set
            {
                group = value;
            }
        }

        private uint numbersOfDisc;
        public uint NumbersOfDisc
        {
            get
            {
                return numbersOfDisc;
            }
            set
            {
                numbersOfDisc = value;
            }
        }

        private List<string> codes;
        public List<string> Codes
        {
            get
            {
                return codes;
            }
            set
            {
                codes = value;
            }
        }
    }




}