using System;
using System.Collections.Generic;
using System.Windows.Forms;
using FileProcessor;
using System.Threading;

namespace DVRP1
{
    public partial class Form1 : Form
    {
        private List<Discipline> disciplines;
        private List<List<Student>> students;
        private List<List<Selection>> selections;
        private List<string> sheetNames;
        private string error;
        private Logger loger;

        public string StringWrapper(string arg)
        {

            string res = "";
            if (arg == null)
                return res;
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
            error = "0";
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
            try
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
            catch (Exception e)
            {

                return null;
            }
        }

        private uint MinMax(string shitstring)
        {
            if (shitstring == "")
            {
                return 0;
            }
            if (shitstring == "Не обмежена" || shitstring == "Необмежена" || shitstring == "необмежена" || shitstring == "не обмежена " || shitstring == "не обмежена")
            {
                return 500;
            }
            if (shitstring == "У межах ліцензійного набору")
            {
                return 500;
            }
            int res;
            if (!Int32.TryParse(shitstring, out res))
            {
                return 0;
            }
            else
                return Convert.ToUInt32(shitstring);
        }

        private void OpenDisciplinesFiles_button(object sender, EventArgs e)
        {
            DisciplinesFiles();
        }

        private void DisciplinesFiles()
        {
            try
            {
                checkBox2.Enabled = false;
                var path = OpenDialog();
                if (path.Length < 1)
                {
                    return;
                }
                FileReader Dop = new FileReader(path);
                progressBarLabel.Text = "Обробляється";

                error = "0.4";
                un_magistr_select = Convert.ToInt32(Dop.GetCellValue("B", 4, 0));
                un_bacalavr_select = Convert.ToInt32(Dop.GetCellValue("B", 5, 0));
                un_magistr_condit_select = (int)(un_magistr_select * Percentchange(Dop.GetCellValue("C", 4, 0)));
                un_bacalavr_condit_select = (int)(un_bacalavr_select * Percentchange(Dop.GetCellValue("C", 5, 0)));

                fac_magistr_select = Convert.ToInt32(Dop.GetCellValue("B", 8, 0));
                fac_bacalavr_select = Convert.ToInt32(Dop.GetCellValue("B", 9, 0));
                fac_magistr_condit_select = (int)(fac_magistr_select * Percentchange(Dop.GetCellValue("C", 8, 0)));
                fac_bacalavr_condit_select = (int)(fac_bacalavr_select * Percentchange(Dop.GetCellValue("C", 9, 0)));

                uint i = 2;
                while (Dop.GetCellValue("D", i, 1) != "")
                {
                    error = "1." + i.ToString();
                    if (checkBox2.Checked)
                        disciplines.Add(new Discipline(
                                                StringWrapper(Dop.GetCellValue("A", i, 1)),
                                                StringWrapper(Dop.GetCellValue("B", i, 1)),
                                                StringWrapper(Dop.GetCellValue("C", i, 1)),
                                                StringWrapper(Dop.GetCellValue("D", i, 1)),
                                                MinMax(Dop.GetCellValue("E", i, 1)),
                                                MinMax(Dop.GetCellValue("F", i, 1)),
                                                Courses(Dop.GetCellValue("G", i, 1)), un_magistr_select, un_magistr_condit_select,
                                                StringWrapper(Dop.GetCellValue("G", i, 1)),
                                                StringWrapper(Dop.GetCellValue("H", i, 1)),
                                                StringWrapper(Dop.GetCellValue("I", i, 1))
                                                ));
                    else
                        disciplines.Add(new Discipline(
                            StringWrapper(Dop.GetCellValue("A", i, 1)),
                            StringWrapper(Dop.GetCellValue("B", i, 1)),
                            StringWrapper(Dop.GetCellValue("C", i, 1)),
                            StringWrapper(Dop.GetCellValue("D", i, 1)),
                            MinMax(Dop.GetCellValue("E", i, 1)),
                            MinMax(Dop.GetCellValue("F", i, 1)),
                            Courses(Dop.GetCellValue("G", i, 1)), un_bacalavr_select, un_bacalavr_condit_select,
                            StringWrapper(Dop.GetCellValue("G", i, 1)),
                            StringWrapper(Dop.GetCellValue("H", i, 1)),
                            StringWrapper(Dop.GetCellValue("I", i, 1))));
                    i++;
                }
                i = 2;
                while (Dop.GetCellValue("D", i, 2) != "")
                {
                    error = "2." + i.ToString();
                    disciplines.Add(new Discipline(
                        Dop.GetCellValue("A", i, 2),
                        Dop.GetCellValue("E", i, (2)),
                        Dop.GetCellValue("C", i, (2)),
                        Dop.GetCellValue("D", i, (2)),
                        MinMax(Dop.GetCellValue("E", i, 2)),
                        MinMax(Dop.GetCellValue("F", i, 2)),
                        Courses(Dop.GetCellValue("G", i, (2))), fac_bacalavr_select, fac_bacalavr_condit_select,
                        StringWrapper(Dop.GetCellValue("G", i, 2)),
                        StringWrapper(Dop.GetCellValue("H", i, 2)),
                        StringWrapper(Dop.GetCellValue("I", i, 2))));
                    //J, K
                    //рапорт I -> J, J -> K
                    i++;
                }
                i = 2;
                while (Dop.GetCellValue("D", i, 3) != "")
                {
                    error = "3." + i.ToString();
                    disciplines.Add(new Discipline(
                        Dop.GetCellValue("A", i, (3)),
                        Dop.GetCellValue("E", i, (3)),
                        Dop.GetCellValue("C", i, (3)),
                        Dop.GetCellValue("D", i, (3)),
                        MinMax(Dop.GetCellValue("E", i, 3)),
                        MinMax(Dop.GetCellValue("F", i, 3)),
                        Courses(Dop.GetCellValue("G", i, (3))), fac_magistr_select, fac_magistr_condit_select,
                        StringWrapper(Dop.GetCellValue("G", i, 3)),
                        StringWrapper(Dop.GetCellValue("H", i, 3)),
                        StringWrapper(Dop.GetCellValue("I", i, 3))));

                    i++;
                }
                Dop.Close();

                button1.Enabled = true;
                checkBox1.Enabled = true;
                progressBarLabel.Text = "Готово";
            }
            catch (Exception exc)
            {
                MessageBox.Show("Something went wrong while reading the file: " + exc.Message, "Error!");
                disciplines.Clear();
                return;
            }
        }

        private void OpenStudentFiles_button(object sender, EventArgs e) //
        {
            try
            {
                progressBarLabel.Text = "Обробляється";
                var opd = new OpenFileDialog();
                opd.Filter = "*.xlsx | *.xlsx";
                opd.Multiselect = true;
                opd.ShowDialog();
                if (opd.FileNames.Length < 1)
                {
                    progressBarLabel.Text = "Готово";
                    return;
                }
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
                        while (reader.GetCellValue("E", currentRow, sheetCounter) != "")
                        {
                            error = sheetCounter.ToString() + '.' + currentRow.ToString();
                            Student st = new Student
                                (
                                reader.FacultyFromFileName(),//faculty
                                reader.GetCellValue("D", currentRow, sheetCounter),//email
                                StringWrapper(reader.GetCellValue("E", currentRow, sheetCounter)),//name
                                StringWrapper(reader.GetCellValue("G", currentRow, sheetCounter)),//group
                                Convert.ToUInt32(reader.GetCellValue("F", currentRow, sheetCounter))//NumOfSelections
                                );
                            for (int j = reader.LetterToInt("H"); j < st.NumberOfSelections + reader.LetterToInt("H"); j++)
                            {
                                st.AddChoice(StringWrapper(reader.GetCellValue(reader.IntToLetter(j), currentRow, sheetCounter)));
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
                    //IllyaValidation(students);
                    //CreateOutputValidStudentsFile();
                    progressBar1.PerformStep();
                }
                progressBar1.Value = 0;

                for (int i = 0; i < sheetNames.Count; i++)
                {
                    ComputeStudentsSelections(students[i], selections[i], i);

                }


                button2.Enabled = true;
                checkBox1.Enabled = false;
                progressBarLabel.Text = "Готово";
            }
            catch (Exception ec)
            {
                MessageBox.Show("Something went wrong while reading the file: " + ec.Message, "Error!");
                return;
            }
        }

        private void ComputeStudentsSelections(List<Student> students, List<Selection> selections, int sheetNum)
        {
            progressBar1.Value = 0;
            progressBar1.Maximum = students.Count;
            progressBar1.Step = 1;
            foreach (var student in students)
            {
                for (int i = 0; i < student.Codes.Count; i++)
                {
                    bool found = false;
                    for (int j = 0; j < selections.Count; j++)
                    {
                        if (selections[j].Discipline.Code == student.Codes[i])
                        {
                            selections[j].Students.Add(student);
                            found = true;
                        }
                    }
                    if (!found && checkBox1.Checked)
                    {
                        loger.Log(student.Name + " " + student.Codes[i] + " " + sheetNames[sheetNum]);
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
            try
            {
                progressBarLabel.Text = "Обробляється";

                var fbd = new FolderBrowserDialog();
                fbd.ShowDialog();
                if (fbd.SelectedPath.Length < 1)
                {
                    progressBarLabel.Text = "Готово";
                    return;
                }
                for (int i = 0; i < sheetNames.Count; i++)
                {
                    CreateOutputFile(selections[i], fbd.SelectedPath + "/output" + sheetNames[i] + ".xlsx", i);
                }
                progressBarLabel.Text = "Готово";
                ReloadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something went wrong while writing the file: " + ex.Message, "Error!");
                return;
            }
        }

        private void CreateOutputFile(List<Selection> selection, string filename, int numb)
        {
            FileWriter writer = new FileWriter(filename);
            writer.InsertWorksheet("обрані");
            writer.SetCellValue("C", 1, "Обрані дисципліни", "обрані");
            writer.SetCellValue("A", 2, "№", "обрані");
            writer.SetCellValue("B", 2, "Шифр", "обрані");
            writer.SetCellValue("C", 2, "Назва дисципліни", "обрані");
            writer.SetCellValue("D", 2, "Факультет", "обрані");
            writer.SetCellValue("E", 2, "Кількість студентів", "обрані");
            writer.SetCellValue("F", 2, "мін.", "обрані");
            writer.SetCellValue("G", 2, "макс.", "обрані");
            uint currentCellNumber = 3;
            for (int i = 0; i < selection.Count; i++)
            {
                if (selection[i].Status == 2)
                {
                    writer.SetCellValue("A", currentCellNumber, (currentCellNumber + 1).ToString(), "обрані");
                    writer.SetCellValue("B", currentCellNumber, selection[i].Discipline.Code, "обрані");
                    writer.SetCellValue("C", currentCellNumber, selection[i].Discipline.Name, "обрані");
                    writer.SetCellValue("D", currentCellNumber, selection[i].Discipline.Faculty, "обрані");
                    writer.SetCellValue("E", currentCellNumber, selection[i].Students.Count.ToString(), "обрані");
                    writer.SetCellValue("F", currentCellNumber, selection[i].Discipline.Min.ToString(), "обрані");
                    writer.SetCellValue("G", currentCellNumber, selection[i].Discipline.Max.ToString(), "обрані");
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
            for (int i = 0; i < selection.Count; i++)
            {
                if (selection[i].Status == 1)
                {
                    writer.SetCellValue("A", currentCellNumber, (currentCellNumber + 1).ToString(), "умовно");
                    writer.SetCellValue("B", currentCellNumber, selection[i].Discipline.Code, "умовно");
                    writer.SetCellValue("C", currentCellNumber, selection[i].Discipline.Name, "умовно");
                    writer.SetCellValue("D", currentCellNumber, selection[i].Discipline.Faculty, "умовно");
                    writer.SetCellValue("E", currentCellNumber, selection[i].Students.Count.ToString(), "умовно");
                    writer.SetCellValue("F", currentCellNumber, selection[i].Discipline.Min.ToString(), "умовно");
                    writer.SetCellValue("G", currentCellNumber, selection[i].Discipline.Max.ToString(), "умовно");
                    currentCellNumber++;
                }
            }

            currentCellNumber = 2;
            int stNum = 1;
            writer.InsertWorksheet("рапорт");

            progressBar1.Maximum = selection.Count;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            for (int i = 0; i < selection.Count; i++)
            {
                if (selection[i].Students.Count < 1)
                {
                    progressBar1.PerformStep();
                    continue;
                }
                writer.SetCellValue("C", currentCellNumber, selection[i].Discipline.Code, "рапорт");
                writer.SetCellValue("D", currentCellNumber, selection[i].Discipline.Name, "рапорт");
                writer.SetCellValue("E", currentCellNumber, selection[i].Discipline.Max.ToString(), "рапорт");
                writer.SetCellValue("F", currentCellNumber, selection[i].Discipline.G, "рапорт");
                writer.SetCellValue("G", currentCellNumber, selection[i].Discipline.H, "рапорт");
                writer.SetCellValue("H", currentCellNumber, selection[i].Discipline.I, "рапорт");

                currentCellNumber++;
                for (int j = 0; j < selection[i].Students.Count; j++)
                {
                    writer.SetCellValue("A", currentCellNumber, stNum++.ToString(), "рапорт");
                    writer.SetCellValue("B", currentCellNumber, (j + 1).ToString(), "рапорт");
                    writer.SetCellValue("C", currentCellNumber, selection[i].Students[j].Name, "рапорт");
                    writer.SetCellValue("D", currentCellNumber, selection[i].Students[j].Group, "рапорт");
                    writer.SetCellValue("E", currentCellNumber, selection[i].Students[j].Email, "рапорт");
                    currentCellNumber++;
                }
                progressBar1.PerformStep();


            }

            currentCellNumber = 2;
            stNum = 1;
            writer.InsertWorksheet("рапорт 2");

            progressBar1.Maximum = selection.Count;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            for (int i = 0; i < selection.Count; i++)
            {
                if (selection[i].Students.Count < 1)
                {
                    progressBar1.PerformStep();
                    continue;
                }
                stNum = 1;
                for (int j = 0; j < selection[i].Students.Count; j++)
                {
                    writer.SetCellValue("A", currentCellNumber, currentCellNumber.ToString(), "рапорт 2");
                    writer.SetCellValue("B", currentCellNumber, stNum++.ToString(), "рапорт 2");
                    writer.SetCellValue("C", currentCellNumber, selection[i].Students[j].Name, "рапорт 2");
                    writer.SetCellValue("D", currentCellNumber, selection[i].Students[j].Group, "рапорт 2");
                    writer.SetCellValue("E", currentCellNumber, selection[i].Discipline.Code, "рапорт 2");
                    writer.SetCellValue("F", currentCellNumber, selection[i].Discipline.Name, "рапорт 2");
                    writer.SetCellValue("G", currentCellNumber, selection[i].Students[j].Email, "рапорт 2");
                    currentCellNumber++;
                }
                progressBar1.PerformStep();
                currentCellNumber++;
            }

            currentCellNumber = 2;
            stNum = 1;
            writer.InsertWorksheet("рапорт 3 (не обрані)");

            progressBar1.Maximum = selection.Count;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            for (int i = 0; i < selection.Count; i++)
            {
                if (!(selection[i].Students.Count > 0 && selection[i].Status == 0))
                {
                    progressBar1.PerformStep();
                    continue;
                }
                stNum = 1;
                for (int j = 0; j < selection[i].Students.Count; j++)
                {
                    writer.SetCellValue("A", currentCellNumber, currentCellNumber.ToString(), "рапорт 3 (не обрані)");
                    writer.SetCellValue("B", currentCellNumber, stNum++.ToString(), "рапорт 3 (не обрані)");
                    writer.SetCellValue("C", currentCellNumber, selection[i].Students[j].Name, "рапорт 3 (не обрані)");
                    writer.SetCellValue("D", currentCellNumber, selection[i].Students[j].Group, "рапорт 3 (не обрані)");
                    writer.SetCellValue("E", currentCellNumber, selection[i].Discipline.Code, "рапорт 3 (не обрані)");
                    writer.SetCellValue("F", currentCellNumber, selection[i].Discipline.Name, "рапорт 3 (не обрані)");
                    currentCellNumber++;
                }
                progressBar1.PerformStep();
                currentCellNumber++;
            }

            progressBar1.Maximum = students.Count;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            writer.InsertWorksheet("студенти");
            currentCellNumber = 1;
            for (int j = 0; j < students[numb].Count; j++)
            {
                writer.SetCellValue("A", currentCellNumber, students[numb][j].Name, "студенти");
                writer.SetCellValue("B", currentCellNumber, students[numb][j].Group, "студенти");
                for (int c = 0; c < students[numb][j].Codes.Count; c++)
                {
                    writer.SetCellValue(writer.IntToLetter(c + 3), currentCellNumber, students[numb][j].Codes[c], "студенти");
                }
                currentCellNumber++;
                progressBar1.PerformStep();
            }


            progressBar1.Maximum = selection.Count;
            progressBar1.Value = 0;
            progressBar1.Step = 1;
            writer.InsertWorksheet("статистика");
            currentCellNumber = 1;
            for (int j = 0; j < selection.Count; j++)
            {
                if (selection[j].Students.Count != 0)
                {
                    writer.SetCellValue("A", currentCellNumber, selection[j].Discipline.Code, "статистика");
                    writer.SetCellValue("B", currentCellNumber, selection[j].Discipline.Name, "статистика");
                    writer.SetCellValue("C", currentCellNumber, selection[j].Students.Count.ToString(), "статистика");
                    currentCellNumber++;
                }
                progressBar1.PerformStep();
            }
            writer.CloseAndExport();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            loger = new Logger("Log.txt");
        }

        private void ReloadData()
        {
            for (int i = 0; i < selections.Count; i++)
            {
                selections[i].Clear();
            }
            selections.Clear();
            for (int i = 0; i < students.Count; i++)
            {
                students[i].Clear();
            }

            students.Clear();
            disciplines.Clear();
            sheetNames.Clear();
            button1.Enabled = false;
            button2.Enabled = false;
            checkBox1.Enabled = false;
            checkBox2.Enabled = true;
        }

        private double Percentchange(string percentstring)
        {
            return Convert.ToDouble(percentstring.Trim(new char[] { '>', '=', '%' })) / 100;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var m = new Help();
            m.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Додаток був розроблений у рамках курсової роботи студентами Лєшановим Андрієм та Тімоєєвою Марією", "Автори");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ReloadData();
        }
    }
}
