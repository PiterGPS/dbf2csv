using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace dbf2csv
{
    
    class dbfToCSV
    {
        public string dbf_filename;
        public string csv_filename;
        
        public  dbfToCSV()
        {
           
        }

        

        public bool Do()
        {
            try{
            string Line = "";
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(dbf_filename) + ";Extended Properties=dBASE IV;";
            System.IO.StreamWriter sw = new StreamWriter(csv_filename, false, Encoding.GetEncoding("windows-1251"));
            OleDbConnection oc = new OleDbConnection(strConn);
            oc.Open();
            
            OleDbCommand command = new OleDbCommand("select * from " + Path.GetFileName(dbf_filename), oc);

            bool write_header = true;

            // Execute the DataReader and access the data.
            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                if (write_header)
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        if (Line != "") Line += "\t";
                        Line += reader.GetName(i);
                    }
                    write_header = false;
                    sw.WriteLine(Line);
                }

                Line = "";
                for (int i = 0; i < reader.FieldCount; i++)
                {
                   
                    string ttn=reader.GetDataTypeName(i);
                    string vp="";
                    if (!reader.IsDBNull(i))
                    {
                        if (ttn == "DBTYPE_DATE" || ttn=="DBTYPE_DBTIMESTAMP")
                        {

                            vp = reader.GetDateTime(i).ToString("dd.MM.yyyy");


                        }

                        if (ttn == "DBTYPE_14" )
                        {
                            vp = reader.GetInt32(i).ToString();
                            vp = vp.Replace(",", ".");

                        }

                        if (ttn == "DBTYPE_BOOL")
                        {
                            bool bl = reader.GetBoolean(i);
                            if (bl)
                            {
                                vp = "true";
                            }else{
                                vp = "false";
                            }

                        }

                        if (ttn == "DBTYPE_DECIMAL")
                        {
                            vp = reader.GetDecimal(i).ToString();
                            vp = vp.Replace(",", ".");

                        }

                        if (ttn == "DBTYPE_R8" || ttn == "DBTYPE_CY")
                        {
                            vp = reader.GetDouble(i).ToString("#.##");
                            vp = vp.Replace(",", ".");
                          
                        }

                        if (ttn == "DBTYPE_STR" || ttn == "DBTYPE_WSTR" || ttn == "DBTYPE_WVARCHAR")
                        {
                            vp = reader.GetString(i);
                        }
                    }
                     if (Line != "") Line += "\t";
                     Line += vp;
                }
                sw.WriteLine(Line);
            }

            // Call Close when done reading.
            sw.Close();
            reader.Close();
            return true;
         }
            catch (Exception ex)
            {
                Console.WriteLine("Error: "+ex.Message);
                return false;
          }
            

        }
      

    }
}
