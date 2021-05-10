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
        List<Discipline> disciplines;
        List<Selection> selections;
        List<Student> students;
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
                return new uint[1] { 0};
            }
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] > 47 && c[i] < 58)
                {
                    a[j] = Convert.ToUInt32(c[i])-48;
                    j++;
                }
                if (c[i] == '-')
                {
                    uint x = a[j-1];
                    x++;
                    while (x != Convert.ToUInt32(c[i+1])-48)
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

        private void StartButton_click(object sender, EventArgs e)
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
            foreach ( Discipline elem in disciplines)
            {
                selections.Add(new Selection(elem));
               
            }
            ;
        } 
        

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
