using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace dbf2csv
{
    class Program
    {
        static void Main(string[] args)
        {
           
            if (args.Length < 1)
            {
                Console.WriteLine(@" This is DBase IV to csv (tab delimiters) simple converter");
                Console.WriteLine(@" generate csv files with ANSI (1251) text");
                Console.WriteLine(@"");
                Console.WriteLine(@" HOW USE IT:");
                Console.WriteLine(@">dbf2csv disk:\inpdir\*.dbf [disk:\outdir]");
                Console.WriteLine(@"Piter Gavrinev 2010 :)");
                Console.WriteLine(@"need .NET 2.0 vesrion");
                Console.ReadKey();
                return;
            }
            Console.WriteLine(@"Cool dbf to csv Piter converter.");
            string inp = args[0];
            string cr = "";
            int mp = inp.LastIndexOf("\\");
            if (mp == -1)
            {
                Console.WriteLine(@"error, need full path to files: disk:path/path/*.dbf");
                return;
            }

                cr = inp.Substring(mp + 1);
                inp = inp.Substring(0, mp);

                string out_directory = inp;
                if (args.Length >= 2)
                {
                    out_directory = args[1];
                }

            string[] files = System.IO.Directory.GetFiles(inp,cr);

            if (!System.IO.Directory.Exists(out_directory))
            {
                System.IO.Directory.CreateDirectory(out_directory);
            }
            Console.WriteLine(@"OutDirectory: " + out_directory);
            dbfToCSV dtc = new dbfToCSV();
            for (int i = 0; i < files.Length; i++)
            {
                string flnm = Path.GetFileNameWithoutExtension(files[i]);
                string out_file = out_directory + "\\" + flnm + ".csv";
                Console.Write("conv: " + files[i]);
                
                dtc.csv_filename = out_file;
                dtc.dbf_filename = files[i];
                if (dtc.Do())  Console.Write(" OK.\r\n");
            }
            
            Console.WriteLine(@"End.");
            
        }
    }
}
