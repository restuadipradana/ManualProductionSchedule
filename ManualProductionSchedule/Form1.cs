using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ManualProductionSchedule
{
    public partial class UploadForm : Form
    {
        public UploadForm()
        {
            InitializeComponent();
        }

        private void UploadForm_Load(object sender, EventArgs e)
        {
            try
            {
                
                string path2 = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string path = @path2 + "\\ManualProductionSchedule";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string[] files = Directory.GetFiles(path, "*.xls*");
                listBox1.Items.AddRange(files);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                button1.Text = "Processing";
                button1.Enabled = false;
                DataTable tmpData = CreateTmpData();
                DataSet dst = new DataSet();
                dst.Tables.Add(tmpData);
                foreach (var idx in listBox1.Items)
                {
                    
                    string fileName = idx.ToString();
                    var f1 = new FileInfo(fileName);
                    List<string> cf = new List<string>();
                    using (var p = new ExcelPackage(f1))
                    {
                        var ak = p.Workbook.View.ActiveTab;
                        var ws = p.Workbook.Worksheets[ak];
                        var endrow = ws.Dimension.End.Row;
                        var endcol = ws.Dimension.End.Column;
                        int start = 0;

                        /*
                        var dtt = ws.Cells[11, 26].Value.ToString();   //<------THIS
                        long dtt3 = Convert.ToInt64(Math.Floor (Convert.ToDouble(ws.Cells[11, 26].Value)));
                        DateTime result3 = DateTime.FromOADate(dtt3);
                        string axab = result3.ToString("MM/dd/yyyy");

                        string axac = DateTime.FromOADate(Convert.ToInt64(Math.Floor(Convert.ToDouble(ws.Cells[57, 26].Value)))).ToString("MM/dd/yyyy"); //a/s
                        cf.Add(axab);
                        string axcv = DateTime.FromOADate(long.Parse(ws.Cells[9, 2].Value.ToString())).ToString("MM/dd/yyyy"); //cfd, pd, crd
                        cf.Add(axcv);
                        DateTime resultx = DateTime.FromOADate(long.Parse(ws.Cells[8, 26].Value.ToString()));
                        long dateNum = long.Parse(ws.Cells[9,1].Value.ToString());
                        DateTime result = DateTime.FromOADate(dateNum);
                        string axaa = result.ToString("MM/dd/yyyy");

                        string abdc = cf[1];
                        cf.Clear();
                        */

                        //int artCnt = 1;
                        string idn = "";
                        string idnx = "";
                        int nt = 0; //nt pos
                        int asw = 1;
                        int awxb = 1;
                        int ends = 0;

                        for (int xtd = 1; xtd <= endrow; xtd++)
                        {
                            asw = awxb;
                            for (int ap = asw; ap <= endrow; ap++)//looking n/t
                            {
                                if (ws.Cells[ap, 26].Value != null)
                                {
                                    if (ws.Cells[ap, 26].Value.Equals("A / S"))
                                    {
                                        nt = ap - 2;
                                        for (int ax = 1; ax < endcol; ax++)
                                        {
                                            if (ws.Cells[nt, ax].Value != null)
                                            {
                                                idnx = ws.Cells[nt, ax].Value.ToString();
                                                idn = idnx.Trim();
                                                
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            start = nt + 4;
                            for (int az = start; az < endrow; az++)
                            {
                                string CFD, PD, CRD, DEST, CUSNO, MODNM, MODEL, ARTICLE, QTY;
                                if (ws.Cells[az, 6].Value == null || ws.Cells[az, 26].Value == null) //, PO, TH
                                {
                                    if (ws.Cells[az + 3, 26].Value != null)
                                    {
                                        if (ws.Cells[az + 3, 26].Value.Equals("TH") || (ws.Cells[az + 3, 26].Value.Equals("Thaønh hình")))
                                        {
                                            awxb = az + 3;
                                            break; //az = az + ...
                                        }
                                        else if (ws.Cells[az + 3, 26].Value.Equals("A / S"))
                                        {
                                            awxb = az + 2;
                                            break;
                                        }
                                        else if (az == endrow - 3)
                                        {
                                            ends = az;
                                            break;
                                        }
                                        else 
                                        {
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        if (az == endrow - 3)
                                        {
                                            ends = az;
                                            break;
                                        }
                                        continue;
                                    }
                                }
                                //else if (ws.Cells[az, 26].Value.Equals("A / S") || ws.Cells[az, 26].Value.Equals("TH"))
                                //{
                                //    awxb = az + 2;
                                //    break;
                                //}

                                if (ws.Cells[az, 1].Value == null) { CFD = " "; }
                                else { CFD = DateTime.FromOADate(long.Parse(ws.Cells[az, 1].Value.ToString())).ToString("MM/dd/yyyy"); }

                                if (ws.Cells[az, 2].Value == null) { PD = " "; }
                                else { PD = DateTime.FromOADate(long.Parse(ws.Cells[az, 2].Value.ToString())).ToString("MM/dd/yyyy"); }

                                if (ws.Cells[az, 3].Value == null) { CRD = " "; }
                                else { CRD = DateTime.FromOADate(long.Parse(ws.Cells[az, 3].Value.ToString())).ToString("MM/dd/yyyy"); }

                                if (ws.Cells[az, 4].Value == null) { DEST = " "; }
                                else { DEST = ws.Cells[az, 4].Value.ToString(); }

                                if (ws.Cells[az, 5].Value == null) { CUSNO = " "; }
                                else { CUSNO = ws.Cells[az, 5].Value.ToString(); }

                                if (ws.Cells[az, 7].Value == null) { MODNM = " "; }
                                else { MODNM = ws.Cells[az, 7].Value.ToString(); }

                                if (ws.Cells[az, 8].Value == null) { MODEL = " "; }
                                else { MODEL = ws.Cells[az, 8].Value.ToString(); }

                                if (ws.Cells[az, 9].Value == null) { ARTICLE = " "; }
                                else { ARTICLE = ws.Cells[az, 9].Value.ToString(); }

                                if (ws.Cells[az, 10].Value == null) { QTY = " "; }
                                else { QTY = ws.Cells[az, 10].Value.ToString(); }



                                cf.Add(idn); // Identity
                                cf.Add(CFD); // CFD
                                cf.Add(PD); // PD
                                cf.Add(CRD); // CRD
                                cf.Add(DEST); // Dest
                                cf.Add(CUSNO); // Cust No
                                cf.Add(ws.Cells[az, 6].Value.ToString()); // PO*
                                cf.Add(MODNM); // Model Name
                                cf.Add(MODEL); // Model
                                cf.Add(ARTICLE); // Article
                                cf.Add(QTY); // Order Qty
                                cf.Add(DateTime.FromOADate(Convert.ToInt64(Math.Floor(Convert.ToDouble(ws.Cells[az, 26].Value)))).ToString("MM/dd/yyyy")); // A/S*

                                DataRow row = tmpData.NewRow();
                                row["Identity"] = cf[0];
                                row["CFD"] = cf[1];
                                row["PD"] = cf[2];
                                row["CRD"] = cf[3];
                                row["Destination"] = cf[4];
                                row["CustNo"] = cf[5];
                                row["Order"] = cf[6];
                                row["ModelName"] = cf[7];
                                row["Model"] = cf[8];
                                row["Article"] = cf[9];
                                row["OrdrQty"] = cf[10];
                                row["AS"] = cf[11];
                                tmpData.Rows.Add(row);
                                cf.Clear();

                                if (ws.Cells[az + 4, 26].Value != null)
                                {
                                    if (ws.Cells[az + 4, 26].Value.Equals("A / S"))
                                    {
                                        awxb = az + 3;
                                        break; //az = az + ...
                                    }
                                }
                            }
                            //end of workbook insert
                            if (ends == endrow - 3)
                            {
                                break; //az = az + ...
                            }
                        }
                        
                    }


                    //tmpData.ToCSV(@"C:\testroom\no use\ManualProductionSchedule.csv");
                }
                string loc = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string floc = @loc + "\\ManualProductionSchedule";
                tmpData.ToCSV(@loc+"\\ManualProductionSchedule.csv");
                MessageBox.Show("Convert Successful", "Done",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                button1.Text = "Convert to CSV";
                button1.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error 2",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private static DataTable CreateTmpData()
        {
            DataTable tmpData = new DataTable("ManualProductionSchedule");
            DataColumn[] cols =
            {
                new DataColumn("Identity", typeof(String)),
                new DataColumn("CFD", typeof(String)),
                new DataColumn("PD", typeof(String)),
                new DataColumn("CRD", typeof(String)),
                new DataColumn("Destination", typeof(String)),
                new DataColumn("CustNo", typeof(String)),
                new DataColumn("Order", typeof(String)),
                new DataColumn("ModelName", typeof(String)),
                new DataColumn("Model", typeof(String)),
                new DataColumn("Article", typeof(String)),
                new DataColumn("OrdrQty", typeof(String)),
                new DataColumn("AS", typeof(String))
            };
            tmpData.Columns.AddRange(cols);
            return tmpData;
        }

    }

    public static class CSVUtility
    {
        public static void ToCSV(this DataTable tmpData, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  
            //for (int i = 0; i < tmpData.Columns.Count; i++)
            //{
            //    sw.Write(tmpData.Columns[i]);
            //    if (i < tmpData.Columns.Count - 1)
            //    {
            //        sw.Write(",");
            //    }
            //}
            //sw.Write(sw.NewLine);
            foreach (DataRow dr in tmpData.Rows)
            {
                for (int i = 0; i < tmpData.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < tmpData.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
    }
}
