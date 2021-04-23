﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DVRP1
{
    class Student
    {
        private string faculty;
        private string email;
        private string name;
        private string course;
        private string group;
        private uint numberOfSelections;
        private List<string> codes;

        public Student(string faculty, string email, string name, string course, string group, uint number, string firstChoice)
        {
            this.faculty = faculty;
            this.email = email;
            this.name = name;
            this.course = course;
            this.group = group;
            this.numberOfSelections = number;
            this.codes.Add(firstChoice);
        }

        public Student(Student st)
        {
            this.faculty = st.faculty;
            this.email = st.email;
            this.name = st.name;
            this.course = st.course;
            this.group = st.group;
            this.numberOfSelections = st.numberOfSelections;
            this.codes = st.codes;
        }


        public void AddChoice(string choice)
        {
            this.codes.Add(choice);
        }

        public string Faculty { get => faculty; set => faculty = value; }
        public string Email { get => email; set => email = value; }
        public string Name { get => name; set => name = value; }
        public string Course { get => course; set => course = value; }
        public string Group { get => group; set => group = value; }
        public uint NumberOfSelections { get => numberOfSelections; set => numberOfSelections = value; }
        public List<string> Codes { get => codes; set => codes = value; }
    }
}
