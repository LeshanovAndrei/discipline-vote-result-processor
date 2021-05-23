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
        private List<Student> students1;
        private List<Student> students2;
        private List<Selection> selections1;
        private List<Selection> selections2;
        private Logger loger;
   
       

        string un_magistr_select;
        string un_bacalavr_select;
        string un_magistr_condit_select;
        string un_bacalavr_condit_select;

        string fac_magistr_select;
        string fac_bacalavr_select;
        string fac_magistr_condit_select;
        string fac_bacalavr_condit_select;


        private int FindStudentInList(Student st, List<Student> students)
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
            selections1 = new List<Selection>();
            selections2 = new List<Selection>();
            students1 = new List<Student>();
            students2 = new List<Student>();
            loger = new Logger("C:\\Users\\30m20\\OneDrive\\Документы\\GitHub\\discipline-vote-result-processor\\DVRP1\\DVRP1\\bin\\Debug\\Log.txt");
            
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

            un_magistr_select = Dop.GetCellValue("B", 4, Dop.NameOfSheet(0));
            un_bacalavr_select = Dop.GetCellValue("B", 5, Dop.NameOfSheet(0));
            un_magistr_condit_select = Dop.GetCellValue("E", 4, Dop.NameOfSheet(0));
            un_bacalavr_condit_select = Dop.GetCellValue("E", 5, Dop.NameOfSheet(0));

            fac_magistr_select = Dop.GetCellValue("B", 8, Dop.NameOfSheet(0));
            fac_bacalavr_select = Dop.GetCellValue("B", 9, Dop.NameOfSheet(0));
            fac_magistr_condit_select = Dop.GetCellValue("E", 8, Dop.NameOfSheet(0));
            fac_bacalavr_condit_select = Dop.GetCellValue("E", 9, Dop.NameOfSheet(0));

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
                    Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(1))), 0));
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
                    Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(2))), 1));

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
                    Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(3))), 1));

                i++;
            }
            Dop.Close();
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
                uint currentRow = 2;
                while (reader.GetCellValue("E", currentRow, reader.NameOfSheet(0)) != null && reader.GetCellValue("E", currentRow, reader.NameOfSheet(0)) != "")
                {
                    Student st = new Student
                        (
                        reader.FacultyFromFileName(),//faculty
                        reader.GetCellValue("D", currentRow, reader.NameOfSheet(0)),//email
                        reader.GetCellValue("E", currentRow, reader.NameOfSheet(0)),//name
                        reader.GetCellValue("G", currentRow, reader.NameOfSheet(0)),//group
                        Convert.ToUInt32(reader.GetCellValue("F", currentRow, reader.NameOfSheet(0)))//NumOfSelections
                        );
                    for (int j = reader.LetterToInt("H"); j < st.NumberOfSelections + reader.LetterToInt("H"); j++)
                    {
                        st.AddChoice(reader.GetCellValue(reader.IntToLetter(j), currentRow, reader.NameOfSheet(0)));
                    }
                    currentRow++;
                    int pos = FindStudentInList(st, students1);
                    if (pos != -1)
                    {
                        students1[pos].AddChoice(st.Codes);
                    }
                    else
                    {
                        students1.Add(st);
                    }
                }

                currentRow = 2;
                while (reader.GetCellValue("E", currentRow, reader.NameOfSheet(1)) != null && reader.GetCellValue("E", currentRow, reader.NameOfSheet(1)) != "")
                {
                    Student st = new Student
                        (
                        reader.FacultyFromFileName(),//faculty
                        reader.GetCellValue("D", currentRow, reader.NameOfSheet(1)),//email
                        reader.GetCellValue("E", currentRow, reader.NameOfSheet(1)),//name
                        reader.GetCellValue("G", currentRow, reader.NameOfSheet(1)),//group
                        Convert.ToUInt32(reader.GetCellValue("F", currentRow, reader.NameOfSheet(1)))//NumOfSelections
                        );
                    for (int j = reader.LetterToInt("H"); j < st.NumberOfSelections + reader.LetterToInt("H"); j++)
                    {
                        st.AddChoice(reader.GetCellValue(reader.IntToLetter(j), currentRow, reader.NameOfSheet(1)));
                    }
                    currentRow++;
                    int pos = FindStudentInList(st, students2);
                    if (pos != -1)
                    {
                        students2[pos].AddChoice(st.Codes);
                    }
                    else
                    {
                        students2.Add(st);
                    }
                }


                reader.Close();
            }

            ComputeStudentsSelections(students1, selections1);
            ComputeStudentsSelections(students2, selections2);
            button2.Enabled = true;
        }

        private void ComputeStudentsSelections(List<Student> students, List<Selection> selections)
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
                    if (disciplines.Exists(x => x.Code == student.Codes[i]))
                    {
                        selections.Add(new Selection(disciplines.Find(x => x.Code == student.Codes[i])));
                        selections.Find(x => x.Discipline.Code == student.Codes[i]).AddStudent(student);
                    }
                    else
                    if (checkBox1.Checked)
                    {
                      
                        loger.Log("Student " + student.Name +" choose wrong discipline: "+ student.Codes[i] );
                    }

                }
            }
        }

        private void CreateOutput(object sender, EventArgs e)
        {
            var fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            CreateOutputFile(selections1, fbd.SelectedPath + "/output1.xlsx");
            CreateOutputFile(selections2, fbd.SelectedPath + "/output2.xlsx");
        }

        private void CreateOutputFile(List<Selection> selections, string filename)
        {
            FileWriter writer = new FileWriter(filename);
            writer.InsertWorksheet("Студенты");//ТУТ БУДУТ ВСЯКИЕ ДРУГИЕ ЛИСТЫ
            uint currentCellNumber = 1;
            for (uint i = 0; i < selections.Count; i++)
            {
                writer.SetCellValue("A", currentCellNumber++, selections[(int)i].Discipline.Name, "Студенты");
                for (uint j = 0; j < selections[(int)i].Students.Count; j++)
                {
                    writer.SetCellValue("B", currentCellNumber++, selections[(int)i].Students[(int)j].Name, "Студенты");
                }
            }
            writer.InsertWorksheet("Количества");
            for (uint i = 0; i < selections.Count; i++)
            {
                writer.SetCellValue("A", i+1, selections[(int)i].Discipline.Name, "Количества");
                writer.SetCellValue("B", i+1, selections[(int)i].Students.Count.ToString(), "Количества");
            }
            MessageBox.Show("Done!");
            writer.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Percentchange(string percentstring)
        {
            double percent = Convert.ToDouble(percentstring.Trim(new char[] { '>', '=', '%' })) / 100;

        }

        
    }
}
