using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using System.IO;
using System.Data;

namespace workkingwithcsv
{
    class CSVHELP
    {
        public bool ReadCSV(string path, out string text)
        {
            text = "\n";
            string[] lines = File.ReadAllLines(path);
            foreach (var line in lines)
            {
                text += line+"\n";
            }
            return true;
        }
        public bool WriteCSV(string[] arr, string path)
        {
            StreamWriter sw = new StreamWriter(path);
            for (int i = 0; i <= arr.Length - 1; i++)
            {
                if (i != arr.Length - 1)
                {
                    sw.Write(arr[i] + ",");
                    sw.Flush();
                }
                else
                {
                    sw.Write(arr[i] + "\n");
                    sw.Flush();
                }
            }
            sw.Close();
            return true;
        }
        public bool InsertCSV(string[] arr, string path)
        {
            StreamWriter sw = File.AppendText(path);
            for (int i = 0; i <= arr.Length-1; i++)
            {
                if (i != arr.Length - 1)
                {
                    sw.Write(arr[i] + ",");
                    sw.Flush();
                }
                else
                {
                    sw.Write(arr[i] + "\n");
                    sw.Flush();
                }
            }
            sw.Close();
            return true;

        }
        public bool UpdateCSV(string[] arr, string path)
        {
          
            string[] lines = File.ReadAllLines(path);
            StreamWriter sw = new StreamWriter(path);
            foreach (var line in lines)
            {
                string[] tests = line.Split(',');
                if (tests[0] == arr[0])
                {
                    for (int i = 0; i <= arr.Length - 1; i++)
                    {
                        if (i != arr.Length - 1)
                        {
                            sw.Write(arr[i] + ",");
                            sw.Flush();
                        }
                        else
                        {
                            sw.Write(arr[i] + "\n");
                            sw.Flush();
                        }
                    }
                }
                else
                {
                    for (int i = 0; i <= tests.Length - 1; i++)
                    {
                        if (i != tests.Length - 1)
                        {
                            sw.Write(tests[i] + ",");
                            sw.Flush();
                        }
                        else
                        {
                            sw.Write(tests[i] + "\n");
                        }
                    }
                }
            }
            sw.Close();
            return true;
        }
        public bool DeleteCSV(string parameter, string path)
        {
           
            string[] limes = File.ReadAllLines(path);
            StreamWriter sw = new StreamWriter(path);
            foreach (var line in limes)
            {
                string[] tests = line.Split(',');
                if (tests[0] != parameter)
                {
                    for (int i = 0; i <= tests.Length - 1; i++)
                    {
                        if (i != tests.Length - 1)
                        {
                            sw.Write(tests[i] + ",");
                            sw.Flush();
                        }
                        else
                        {
                            sw.Write(tests[i] + "\n");
                        }
                    }
                }
            }
            sw.Close();
            return true;
        }
       
        public DataTable ReadCsvToDatatable(string path)
        {
            DataTable dt = new DataTable();
            using (StreamReader rd = new StreamReader(path))
            {
                string[] Headers = rd.ReadLine().Split(',');
                foreach (var header in Headers)
                {
                    dt.Columns.Add(header);
                }
                while (!rd.EndOfStream)
                {
                    string[] read = rd.ReadLine().Split(',');
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i <= Headers.Length - 1; i++)
                    {
                        dr[i] = read[i];
                    }
                    dt.Rows.Add(dr);

                }
            }
            return dt;
        }
        public bool ExportCSVto2GridVew(string path,DataGridView grv1,DataGridView grv2)
        {
            DataTable dt1=new DataTable();
            DataTable dt2=new DataTable();
            using (StreamReader rd = new StreamReader(path))
            {
                string[] heads = rd.ReadLine().Split(',');
                if (heads.Length != 0)
                {
                    foreach (var head in heads)
                    {
                        dt1.Columns.Add(head);
                    }
                }
                
                while (!rd.EndOfStream)
                {
                    string[] reads = rd.ReadLine().Split(',');

                    if (reads.Length>1)
                    {
                        if (reads.Length == heads.Length)
                        {
                            DataRow dr = dt1.NewRow();
                            for (int i = 0; i <= reads.Length - 1; i++)
                            {
                                dr[i] = reads[i];
                            }
                            dt1.Rows.Add(dr);
                        }
                        else
                        {
                            if (dt2.Columns.Count == 0)
                            {
                                foreach (var read in reads)
                                {
                                    dt2.Columns.Add(read);
                                }
                            }
                            else
                            {
                                DataRow dr = dt2.NewRow();
                                for (int i = 0; i <= reads.Length - 1; i++)
                                {
                                    dr[i] = reads[i];
                                }
                                dt2.Rows.Add(dr);
                            }
                        }
                    }
                   
                   
                }
            }
            grv1.DataSource = dt1;
            grv2.DataSource = dt2;
            return true;
        }
        public bool DataGridViewToCsv(DataGridView grv, string Path)
        {
            try
            {
                System.Text.Encoding sjisEnc = System.Text.Encoding.GetEncoding("Shift_JIS");
                StreamWriter sw = new StreamWriter(Path,false,sjisEnc);

                for (int i = 0; i <= grv.Columns.Count - 1; i++)
                {
                    string s = grv.Columns[i].HeaderText;
                    if (i != grv.Columns.Count - 1)
                    {
                        sw.Write(s + ",");
                    }
                    else
                    {
                        sw.Write(s + "\n");
                    }
                }
                for (int i = 0; i <= grv.Rows.Count - 2; i++)
                {

                    for (int j = 0; j <= grv.Rows[0].Cells.Count - 1; j++)
                    {
                        if (j != grv.Rows[0].Cells.Count - 1)
                        {
                            sw.Write(grv.Rows[i].Cells[j].Value.ToString() + ",");

                        }
                        else
                        {
                            sw.Write(grv.Rows[i].Cells[j].Value.ToString());
                        }

                    }
                    sw.Write("\n");
                }

                sw.Close();
            }
            catch (Exception  ex)
            {

                Console.WriteLine(ex.Message);
            }
           
            return true;
        }
        public bool DataGridViewApendCsv(DataGridView grv, string path)
        {
            StreamWriter sw = File.AppendText(path);
            for (int i = 0; i <= grv.Rows.Count - 2; i++)
            {

                for (int j = 0; j <= grv.Rows[0].Cells.Count - 1; j++)
                {
                    if (j != grv.Rows[0].Cells.Count - 1)
                    {
                        sw.Write(grv.Rows[i].Cells[j].Value.ToString() + ",");

                    }
                    else
                    {
                        sw.Write(grv.Rows[i].Cells[j].Value.ToString());
                    }

                }
                sw.Write("\n");
            }

            sw.Close();
            return true;
        }
        public bool DataGridViewToCsvApendWithHeader(DataGridView grv, string Path)
        {

            StreamWriter sw = File.AppendText(Path);

            for (int i = 0; i <= grv.Columns.Count - 1; i++)
            {
                string s = grv.Columns[i].HeaderText;
                if (i != grv.Columns.Count - 1)
                {
                    sw.Write(s + ",");
                }
                else
                {
                    sw.Write(s + "\n");
                }
            }
            for (int i = 0; i <= grv.Rows.Count - 2; i++)
            {

                for (int j = 0; j <= grv.Rows[0].Cells.Count - 1; j++)
                {
                    if (j != grv.Rows[0].Cells.Count - 1)
                    {
                        sw.Write(grv.Rows[i].Cells[j].Value.ToString() + ",");

                    }
                    else
                    {
                        sw.Write(grv.Rows[i].Cells[j].Value.ToString());
                    }

                }
                sw.Write("\n");
            }

            sw.Close();
            return true;
        }
        public bool lstObjtocsv(string[][] data,int rows,int colums, string path)
        {
            StreamWriter sw = File.AppendText(path);
            for (int i = 0; i<= rows-1; i++)
            {
                for (int j = 0; j <= colums-1; j++)
                {
                    sw.Write(data[i][j]);
                    if (j < colums-1)
                        sw.Write(",");
                }
                if (i < rows - 1)
                    sw.Write("\n");
            }
            return true;
        }
       
    }
}
