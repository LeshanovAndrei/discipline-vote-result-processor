using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace DVRP1
{
    class Logger
    {
        public Logger(string filepath)
        {
            this.filepath = filepath;
           
        }
        ~Logger(){

        }
       
        public string filepath;
        public string Filepath
        {
            get
            {
                return filepath;
            }
            set
            {
                filepath = value;
            }
        }
        public void Log(string error)
        {
            using (StreamWriter sw = new StreamWriter(filepath, true, System.Text.Encoding.Default))
            {
                sw.WriteLine(error);
            }

           /* using (FileStream fstream = new FileStream(filepath, FileMode.OpenOrCreate))
            {
               
                byte[] array = System.Text.Encoding.Default.GetBytes(error);
             
                fstream.Write(array, 0, array.Length);
            
            }*/
        }
    }
}
