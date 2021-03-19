using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DVRP1
{
    class Discipline
    {
        public Discipline(string faculty, string cathedra, string code, string name, uint max, uint min, uint[] courses)
        {
            this.faculty = faculty;
            this.cathedra = cathedra;
            this.code = code;
            this.name = name;
            this.max = max;
            this.min = min;
            this.courses = courses;
            
        }
        ~Discipline()
        {
            
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
        private string cathedra;
        public string Cathedra
        {
            get
            {
                return cathedra;
            }
            set
            {
                cathedra = value;
            }
        }
        private string code;
        public string Code
        {
            get
            {
                return code;
            }
            set
            {
                code = value;
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
        private uint max;
        public uint Max
        {
            get
            {
                return max;
            }
            set
            {
                max = value;
            }
        }
        private uint min;
        public uint Min
        {
            get
            {
                return min;
            }
            set
            {
                min = value;
            }
        }
        private uint[] courses;
        public uint[] Courses
        {
            get
            {
                return courses;
            }
            set
            {
                courses = value;
            }
        }

    }
}
