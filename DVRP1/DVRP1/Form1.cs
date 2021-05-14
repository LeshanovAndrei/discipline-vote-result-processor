using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FileProcessor;

namespace DVRP1
{
    public partial class Form1 : Form
    {
        private List<Discipline> disciplines;
        private List<Student> students;
        private List<Selection> selections;

        private int FindStudentInList(Student st)
        {
            for (int i = 0; i < students.Count; i++)
            {
                if (st.Name == students[i].Name)
                {
                    if (st.Group == students[i].Group)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        public Form1()
        {

            InitializeComponent();
            disciplines = new List<Discipline>();
            selections = new List<Selection>();
            students = new List<Student>();
        }
        private string OpenDialog()
        {
            var opd = new OpenFileDialog();
            opd.Filter = "*.xlsx | *.xlsx";
            opd.ShowDialog();
            return opd.FileName;
        }

        private uint[] Courses(string c)
        {

            int j = 0;
            uint[] a = new uint[6];
            if (c == "" || c == null)
            {
                return new uint[1] { 0 };
            }
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] > 47 && c[i] < 58)
                {
                    a[j] = Convert.ToUInt32(c[i]) - 48;
                    j++;
                }
                if (c[i] == '-')
                {
                    uint x = a[j - 1];
                    x++;
                    while (x != Convert.ToUInt32(c[i + 1]) - 48)
                    {
                        a[j] = x;
                        x++;
                        j++;
                    }
                }
            }
            uint[] res = new uint[j];
            for (int i = 0; i < j; i++)
            {
                res[i] = a[i];
            }
            return res;
        }

        private uint MinMaxShitHappens(string shitstring)
        {
            if (shitstring == "")
            {
                return 0;
            }
            if (shitstring == "Не обмежена" || shitstring == "Необмежена" || shitstring == "необмежена" || shitstring == "не обмежена " || shitstring == "не обмежена")
            {
                return 4294967294;
            }
            else
                return Convert.ToUInt32(shitstring);
        }

        public uint[] Func(string str)
        {


            str = str.Replace(",-;", " ");

            char[] array = str.ToCharArray();
            uint[] Uint = new uint[array.Length];
            for (uint i = 0; i < array.Length; i++)
            {
                Uint[i] = Convert.ToUInt32(array[i]);

            }
            return Uint;
        }

        private void OpenDisciplinesFiles_button(object sender, EventArgs e)
        {
            FileReader Dop = new FileReader(OpenDialog());

            string un_magistr_select = Dop.GetCellValue("B", 4, Dop.NameOfSheet(0));
            string un_bacalavr_select = Dop.GetCellValue("B", 5, Dop.NameOfSheet(0));
            string un_magistr_condit_select = Dop.GetCellValue("E", 4, Dop.NameOfSheet(0));
            string un_bacalavr_condit_select = Dop.GetCellValue("E", 5, Dop.NameOfSheet(0));

            string fac_magistr_select = Dop.GetCellValue("B", 8, Dop.NameOfSheet(0));
            string fac_bacalavr_select = Dop.GetCellValue("B", 9, Dop.NameOfSheet(0));
            string fac_magistr_condit_select = Dop.GetCellValue("E", 8, Dop.NameOfSheet(0));
            string fac_bacalavr_condit_select = Dop.GetCellValue("E", 9, Dop.NameOfSheet(0));

            uint i = 2;
            while (Dop.GetCellValue("A", i, Dop.NameOfSheet(1)) != null)
            {
                disciplines.Add(new Discipline(
                    Dop.GetCellValue("A", i, Dop.NameOfSheet(1)),
                    Dop.GetCellValue("B", i, Dop.NameOfSheet(1)),
                    Dop.GetCellValue("C", i, Dop.NameOfSheet(1)),
                    Dop.GetCellValue("D", i, Dop.NameOfSheet(1)),
                    MinMaxShitHappens(Dop.GetCellValue("E", i, Dop.NameOfSheet(1))),
                    MinMaxShitHappens(Dop.GetCellValue("F", i, Dop.NameOfSheet(1))),
                    Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(1)))));
                i++;
            }
            while (Dop.GetCellValue("A", i, Dop.NameOfSheet(2)) != null)
            {
                disciplines.Add(new Discipline(
                    Dop.GetCellValue("A", i, Dop.NameOfSheet(2)),
                    Dop.GetCellValue("E", i, Dop.NameOfSheet(2)),
                    Dop.GetCellValue("C", i, Dop.NameOfSheet(2)),
                    Dop.GetCellValue("D", i, Dop.NameOfSheet(2)),
                    MinMaxShitHappens(Dop.GetCellValue("F", i, Dop.NameOfSheet(2))),
                    0,
                    Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(2)))));

                i++;
            }
            while (Dop.GetCellValue("A", i, Dop.NameOfSheet(3)) != null)
            {
                disciplines.Add(new Discipline(
                    Dop.GetCellValue("A", i, Dop.NameOfSheet(3)),
                    Dop.GetCellValue("E", i, Dop.NameOfSheet(3)),
                    Dop.GetCellValue("C", i, Dop.NameOfSheet(3)),
                    Dop.GetCellValue("D", i, Dop.NameOfSheet(3)),
                    MinMaxShitHappens(Dop.GetCellValue("F", i, Dop.NameOfSheet(3))),
                    0,
                    Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(3)))));

                i++;
            }
            Dop.Close();
            foreach (Discipline elem in disciplines)
            {
                selections.Add(new Selection(elem));

            }
            button1.Enabled = true;
        }


        private void OpenStudentFiles_button(object sender, EventArgs e) //
        {
            var opd = new OpenFileDialog();
            opd.Filter = "*.xlsx | *.xlsx";
            opd.Multiselect = true;
            opd.ShowDialog();
            string[] adresses = opd.FileNames;
            for (int i = 0; i < adresses.Length; i++)
            {
                FileReader reader = new FileReader(adresses[i]);
                uint c = 2;
                for (int listCounter = 0; listCounter < 2; listCounter++)
                {


                    while (reader.GetCellValue("E", c, reader.NameOfSheet(listCounter)) != null && reader.GetCellValue("E", c, reader.NameOfSheet(listCounter)) != "")
                    {
                        Student st = new Student
                            (
                            reader.FacultyFromFileName(),//faculty
                            reader.GetCellValue("D", c, reader.NameOfSheet(listCounter)),//email
                            reader.GetCellValue("E", c, reader.NameOfSheet(listCounter)),//name
                            reader.GetCellValue("G", c, reader.NameOfSheet(listCounter)),//group
                            Convert.ToUInt32(reader.GetCellValue("F", c, reader.NameOfSheet(listCounter)))//NumOfSelections
                            );
                        for (int j = reader.LetterToInt("H"); j < st.NumberOfSelections + reader.LetterToInt("H"); j++)
                        {
                            st.AddChoice(reader.GetCellValue(reader.IntToLetter(j), c, reader.NameOfSheet(listCounter)));
                        }
                        c++;
                        int pos = FindStudentInList(st);
                        if (pos != -1)
                        {
                            students[pos].AddChoice(st.Codes);
                        }
                        else
                        {
                            students.Add(st);
                        }
                    }
                }
                reader.Close();
            }

            ComputeStudentsSelections();
            button2.Enabled = true;
        }

        private void ComputeStudentsSelections()
        {
            foreach (var student in students)
            {
                for (int i = 0; i < student.NumberOfSelections; i++)
                {
                    if (selections.Exists(x => x.Discipline.Code == student.Codes[i]))
                    {
                        selections.Find(x => x.Discipline.Code == student.Codes[i]).AddStudent(student);

                    }
                    else
                    if (checkBox1.Checked)
                        MessageBox.Show("Student " + student.Name + " choose wrong discipline!");

                }
            }
        }

        private void CreatePutput(object sender, EventArgs e)
        {
            //var fbd = new FolderBrowserDialog();
            //DialogResult result = fbd.ShowDialog();
            FileWriter writer = new FileWriter("output.xlsx");
            writer.InsertWorksheet("0");
            uint currentCellNumber = 1;
            for (uint i = 0; i < selections.Count; i++)
            {
                writer.SetCellValue("A", currentCellNumber++, selections[(int)i].Discipline.Name, "0");
                for (uint j = 0; j < selections[(int)i].Students.Count; j++)
                {
                    writer.SetCellValue("B", currentCellNumber++, selections[(int)i].Students[(int)j].Name, "0");
                }
            }
            MessageBox.Show("Done!");
            writer.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
