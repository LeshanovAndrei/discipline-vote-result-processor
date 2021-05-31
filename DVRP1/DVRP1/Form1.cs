using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using FileProcessor;

namespace DVRP1
{
    public partial class Form1 : Form
    {
        private List<Discipline> disciplines;
        private List<List<Student>> students;
        private List<List<Selection>> selections;
        private List<string> sheetNames;
        private Logger loger;

        public string StringWrapper(string arg)
        {
            string res = "";
            for (int i = 0; i < arg.Length; i++)
            {
                if (arg[i] == 'y')
                {
                    res += 'у';
                    continue;
                }
                if (arg[i] == 'c')
                {
                    res += 'с';
                    continue;
                }
                if (arg[i] == 'o')
                {
                    res += 'о';
                    continue;
                }
                if (arg[i] == 'a')
                {
                    res += 'а';
                    continue;
                }
                res += arg[i];
            }
            return res.Trim();
        }

        int un_magistr_select;
        int un_bacalavr_select;
        int un_magistr_condit_select;
        int un_bacalavr_condit_select;

        int fac_magistr_select;
        int fac_bacalavr_select;
        int fac_magistr_condit_select;
        int fac_bacalavr_condit_select;


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
            selections = new List<List<Selection>>();
            students = new List<List<Student>>();
            sheetNames = new List<string>();

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
            if (shitstring == "У межах ліцензійного набору")
            {
                return 4294967294;
            }
            else
                return Convert.ToUInt32(shitstring);
        }

        private void OpenDisciplinesFiles_button(object sender, EventArgs e)
        {
            try
            {
                var path = OpenDialog();
                if (path.Length < 1)
                {
                    return;
                }
                FileReader Dop = new FileReader(path);
                progressBarLabel.Text = "Обробляється";


                un_magistr_select = Convert.ToInt32(Dop.GetCellValue("B", 4, Dop.NameOfSheet(0)));
                un_bacalavr_select = Convert.ToInt32(Dop.GetCellValue("B", 5, Dop.NameOfSheet(0)));
                un_magistr_condit_select = (int)(un_bacalavr_select * Percentchange(Dop.GetCellValue("C", 4, Dop.NameOfSheet(0))));
                un_bacalavr_condit_select = (int)(un_bacalavr_select * Percentchange(Dop.GetCellValue("C", 5, Dop.NameOfSheet(0))));

                fac_magistr_select = Convert.ToInt32(Dop.GetCellValue("B", 8, Dop.NameOfSheet(0)));
                fac_bacalavr_select = Convert.ToInt32(Dop.GetCellValue("B", 9, Dop.NameOfSheet(0)));
                fac_magistr_condit_select = (int)(un_bacalavr_select * Percentchange(Dop.GetCellValue("C", 8, Dop.NameOfSheet(0))));
                fac_bacalavr_condit_select = (int)(un_bacalavr_select * Percentchange(Dop.GetCellValue("C", 9, Dop.NameOfSheet(0))));

                uint i = 2;
                while (Dop.GetCellValue("A", i, Dop.NameOfSheet(1)) != null)
                {
                    disciplines.Add(new Discipline(
                        StringWrapper(Dop.GetCellValue("A", i, Dop.NameOfSheet(1))),
                        StringWrapper(Dop.GetCellValue("B", i, Dop.NameOfSheet(1))),
                        StringWrapper(Dop.GetCellValue("C", i, Dop.NameOfSheet(1))),
                        StringWrapper(Dop.GetCellValue("D", i, Dop.NameOfSheet(1))),
                        MinMaxShitHappens(Dop.GetCellValue("E", i, Dop.NameOfSheet(1))),
                        MinMaxShitHappens(Dop.GetCellValue("F", i, Dop.NameOfSheet(1))),
                        Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(1))), un_bacalavr_select, un_bacalavr_condit_select));
                    i++;
                }
                i = 2;
                while (Dop.GetCellValue("A", i, Dop.NameOfSheet(2)) != null)
                {
                    disciplines.Add(new Discipline(
                        Dop.GetCellValue("A", i, Dop.NameOfSheet(2)),
                        Dop.GetCellValue("E", i, Dop.NameOfSheet(2)),
                        Dop.GetCellValue("C", i, Dop.NameOfSheet(2)),
                        Dop.GetCellValue("D", i, Dop.NameOfSheet(2)),
                        MinMaxShitHappens(Dop.GetCellValue("F", i, Dop.NameOfSheet(2))),
                        0,
                        Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(2))), fac_bacalavr_select, fac_bacalavr_condit_select));

                    i++;
                }
                i = 2;
                while (Dop.GetCellValue("A", i, Dop.NameOfSheet(3)) != null)
                {
                    disciplines.Add(new Discipline(
                        Dop.GetCellValue("A", i, Dop.NameOfSheet(3)),
                        Dop.GetCellValue("E", i, Dop.NameOfSheet(3)),
                        Dop.GetCellValue("C", i, Dop.NameOfSheet(3)),
                        Dop.GetCellValue("D", i, Dop.NameOfSheet(3)),
                        MinMaxShitHappens(Dop.GetCellValue("F", i, Dop.NameOfSheet(3))),
                        0,
                        Courses(Dop.GetCellValue("G", i, Dop.NameOfSheet(3))), fac_magistr_select, fac_magistr_condit_select));

                    i++;
                }
                Dop.Close();

                button1.Enabled = true;
                progressBarLabel.Text = "Готово";
            }
            catch (Exception)
            {
                MessageBox.Show("Something went wrong while reading the file! Try again.", "Error!");
                return;
            }



        }


        private void OpenStudentFiles_button(object sender, EventArgs e) //
        {
            Stopwatch stopwatch = new Stopwatch();
            try
            {
                progressBarLabel.Text = "Обробляється";
                var opd = new OpenFileDialog();
                opd.Filter = "*.xlsx | *.xlsx";
                opd.Multiselect = true;
                opd.ShowDialog();
                stopwatch.Start();
                string[] adresses = opd.FileNames;
                progressBar1.Value = 0;
                progressBar1.Maximum = adresses.Length;
                progressBar1.Step = 1;
                //files
                for (int i = 0; i < adresses.Length; i++)
                {
                    FileReader reader = new FileReader(adresses[i]);
                    uint currentRow = 2;
                    //sheets
                    if (sheetNames.Count == 1 && sheetNames[0] == "маг.")
                        for (int c = 0; c < disciplines.Count; c++)
                        {
                            disciplines[i].ForChoose = un_magistr_select;
                            disciplines[i].ForConditChoose = un_magistr_condit_select;
                        }
                    for (int sheetCounter = 0; sheetCounter < reader.SheetNumber(); sheetCounter++)
                    {
                        currentRow = 2;
                        if (!sheetNames.Contains(reader.NameOfSheet(sheetCounter)))
                        {


                            selections.Add(new List<Selection>());
                            for (int j = 0; j < disciplines.Count; j++)
                            {
                                selections[sheetCounter].Add(new Selection(disciplines[j]));
                            }
                            students.Add(new List<Student>());
                            sheetNames.Add(reader.NameOfSheet(sheetCounter));

                        }
                        //students
                        while (reader.GetCellValue("E", currentRow, reader.NameOfSheet(sheetCounter)) != null && reader.GetCellValue("E", currentRow, reader.NameOfSheet(sheetCounter)) != "")
                        {
                            Student st = new Student
                                (
                                reader.FacultyFromFileName(),//faculty
                                reader.GetCellValue("D", currentRow, reader.NameOfSheet(sheetCounter)),//email
                                reader.GetCellValue("E", currentRow, reader.NameOfSheet(sheetCounter)),//name
                                StringWrapper(reader.GetCellValue("G", currentRow, reader.NameOfSheet(sheetCounter))),//group
                                Convert.ToUInt32(reader.GetCellValue("F", currentRow, reader.NameOfSheet(sheetCounter)))//NumOfSelections
                                );
                            for (int j = reader.LetterToInt("H"); j < st.NumberOfSelections + reader.LetterToInt("H"); j++)
                            {
                                st.AddChoice(StringWrapper(reader.GetCellValue(reader.IntToLetter(j), currentRow, reader.NameOfSheet(sheetCounter))));
                            }
                            currentRow++;
                            int pos = FindStudentInList(st, students[sheetCounter]);
                            if (pos != -1)
                            {
                                students[sheetCounter][pos].AddChoice(st.Codes);
                            }
                            else
                            {
                                students[sheetCounter].Add(st);
                            }
                        }
                    }
                    reader.Close();
                    progressBar1.PerformStep();
                }
                progressBar1.Value = 0;

                for (int i = 0; i < sheetNames.Count; i++)
                {
                    ComputeStudentsSelections(students[i], selections[i]);

                }


                button2.Enabled = true;
                progressBarLabel.Text = "Готово";
                stopwatch.Stop();
                MessageBox.Show(stopwatch.ElapsedMilliseconds.ToString());
            }
            catch (Exception)
            {
                MessageBox.Show("Something went wrong while reading the file! Try again.", "Error!");
                return;
            }
        }

        private void ComputeStudentsSelections(List<Student> students, List<Selection> selections)
        {
            progressBar1.Value = 0;
            progressBar1.Maximum = students.Count;
            progressBar1.Step = 1;
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

                        loger.Log("Student " + student.Name + " choose wrong discipline: " + student.Codes[i]);
                    }

                }
                progressBar1.PerformStep();
            }
            progressBar1.Value = 0;
            progressBar1.Maximum = selections.Count;
            progressBar1.Step = 1;
            for (int i = 0; i < selections.Count; i++)
            {
                selections[i].StatusCheck();
                progressBar1.PerformStep();
            }
        }

        private void CreateOutput(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch();
            
            progressBarLabel.Text = "Обробляється";

            var fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            stopwatch.Start();
            for (int i = 0; i < sheetNames.Count; i++)
            {
                CreateOutputFile(selections[i], fbd.SelectedPath + "/output" + sheetNames[i] + ".xlsx");
            }
            progressBarLabel.Text = "Готово";
            stopwatch.Stop();
            MessageBox.Show(students[0].Count.ToString());
            MessageBox.Show(stopwatch.ElapsedMilliseconds.ToString());
            selections.Clear();
            students.Clear();
            disciplines.Clear();
            button1.Enabled = false;
            button2.Enabled = false;
        }

        private void CreateOutputFile(List<Selection> selections, string filename)
        {
            FileWriter writer = new FileWriter(filename);
            writer.InsertWorksheet("обрані");//ТУТ БУДУТ ВСЯКИЕ ДРУГИЕ ЛИСТЫ
            writer.SetCellValue("C", 1, "Обрані дисципліни", "обрані");
            writer.SetCellValue("A", 2, "№", "обрані");
            writer.SetCellValue("B", 2, "Шифр", "обрані");
            writer.SetCellValue("C", 2, "Назва дисципліни", "обрані");
            writer.SetCellValue("D", 2, "Факультет", "обрані");
            writer.SetCellValue("E", 2, "Кількість студентів", "обрані");
            writer.SetCellValue("F", 2, "мін.", "обрані");
            writer.SetCellValue("G", 2, "макс.", "обрані");
            uint currentCellNumber = 3;
            for (int i = 0; i < selections.Count; i++)
            {
                if (selections[i].Status == 2)
                {
                    writer.SetCellValue("A", currentCellNumber, (currentCellNumber + 1).ToString(), "обрані");
                    writer.SetCellValue("B", currentCellNumber, selections[i].Discipline.Code, "обрані");
                    writer.SetCellValue("C", currentCellNumber, selections[i].Discipline.Name, "обрані");
                    writer.SetCellValue("D", currentCellNumber, selections[i].Discipline.Faculty, "обрані");
                    writer.SetCellValue("E", currentCellNumber, selections[i].Students.Count.ToString(), "обрані");
                    writer.SetCellValue("F", currentCellNumber, selections[i].Discipline.Min.ToString(), "обрані");
                    writer.SetCellValue("G", currentCellNumber, selections[i].Discipline.Max.ToString(), "обрані");
                    currentCellNumber++;
                }



            }

            currentCellNumber = 3;
            writer.InsertWorksheet("умовно");
            writer.SetCellValue("C", 2, "умовно Обрані дисципліни", "умовно");
            writer.SetCellValue("A", 2, "№", "умовно");
            writer.SetCellValue("B", 2, "Шифр", "умовно");
            writer.SetCellValue("C", 2, "Назва дисципліни", "умовно");
            writer.SetCellValue("D", 2, "Факультет", "умовно");
            writer.SetCellValue("E", 2, "Кількість студентів", "умовно");
            writer.SetCellValue("F", 2, "мін.", "умовно");
            writer.SetCellValue("G", 2, "макс.", "умовно");
            for (int i = 0; i < selections.Count; i++)
            {
                if (selections[i].Status == 1)
                {
                    writer.SetCellValue("A", currentCellNumber, (currentCellNumber + 1).ToString(), "умовно");
                    writer.SetCellValue("B", currentCellNumber, selections[i].Discipline.Code, "умовно");
                    writer.SetCellValue("C", currentCellNumber, selections[i].Discipline.Name, "умовно");
                    writer.SetCellValue("D", currentCellNumber, selections[i].Discipline.Faculty, "умовно");
                    writer.SetCellValue("E", currentCellNumber, selections[i].Students.Count.ToString(), "умовно");
                    writer.SetCellValue("F", currentCellNumber, selections[i].Discipline.Min.ToString(), "умовно");
                    writer.SetCellValue("G", currentCellNumber, selections[i].Discipline.Max.ToString(), "умовно");
                    currentCellNumber++;
                }
            }

            currentCellNumber = 2;
            int stNum = 1;
            writer.InsertWorksheet("рапорт");

            progressBar1.Maximum = selections.Count;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            for (int i = 0; i < selections.Count; i++)
            {
                if (selections[i].Students.Count < 1)
                {
                    progressBar1.PerformStep();
                    continue;
                }
                writer.SetCellValue("C", currentCellNumber, selections[i].Discipline.Code, "рапорт");
                writer.SetCellValue("D", currentCellNumber, selections[i].Discipline.Name, "рапорт");
                currentCellNumber++;
                for (int j = 0; j < selections[i].Students.Count; j++)
                {
                    writer.SetCellValue("A", currentCellNumber, stNum++.ToString(), "рапорт");
                    writer.SetCellValue("B", currentCellNumber, (j + 1).ToString(), "рапорт");
                    writer.SetCellValue("C", currentCellNumber, selections[i].Students[j].Name, "рапорт");
                    writer.SetCellValue("D", currentCellNumber, selections[i].Students[j].Group, "рапорт");
                    currentCellNumber++;
                    //stNum++;
                }
                progressBar1.PerformStep();

            }
            writer.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            loger = new Logger("Log.txt");
        }

        private double Percentchange(string percentstring)
        {
            return Convert.ToDouble(percentstring.Trim(new char[] { '>', '=', '%' })) / 100;

        }

    }
}
