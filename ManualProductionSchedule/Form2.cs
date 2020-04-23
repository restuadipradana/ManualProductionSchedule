using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Reflection;
using System.Windows.Forms;

namespace ManualProductionSchedule
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public int Stat;
        struct FtpSetting
        {
            public string Server { get; set; }
            public string Username { get; set; }
            public string Password { get; set; }
            public string FileName { get; set; }
            public string FullName { get; set; }
        }

        FtpSetting _inputParameter;
        private void button1_Click(object sender, EventArgs e)
        {
            string uid, pwd;
            uid = textBox1.Text;
            pwd = textBox2.Text;
            string loc = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            _inputParameter.Username = textBox1.Text;
            _inputParameter.Password = textBox2.Text;
            _inputParameter.Server = "ftp://10.1.0.58";
            _inputParameter.FileName = "ie383.csv";
            _inputParameter.FullName = loc+ "\\ie383.csv";
            backgroundWorker.RunWorkerAsync(_inputParameter);
            button1.Enabled = false;
            button2.Enabled = false;
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string fileName = ((FtpSetting)e.Argument).FileName;
            string fullName = ((FtpSetting)e.Argument).FullName;
            string userName = ((FtpSetting)e.Argument).Username;
            string password = ((FtpSetting)e.Argument).Password;
            string server = ((FtpSetting)e.Argument).Server;
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(new Uri(string.Format("{0}/{1}", server, fileName)));
            request.Method = WebRequestMethods.Ftp.UploadFile;
            request.Credentials = new NetworkCredential(userName, password);
            Stream ftpStream = request.GetRequestStream();
            FileStream fs = File.OpenRead(fullName);
            byte[] buffer = new byte[1024];
            double total = (double)fs.Length;
            int byteRead = 0;
            double read = 0;
            do
            {
                if (!backgroundWorker.CancellationPending)
                {
                    //Upload file & update process bar
                    byteRead = fs.Read(buffer, 0, 1024);
                    ftpStream.Write(buffer, 0, byteRead);
                    read += (double)byteRead;
                    double percentage = read / total * 100;
                    backgroundWorker.ReportProgress((int)percentage);
                    Stat = Convert.ToInt32(percentage);
                }
            }
            while (byteRead != 0);
            
            fs.Close();
            ftpStream.Close();
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            lblStatus.Text = $"{e.ProgressPercentage}% Uploaded ";
            progressBar.Value = e.ProgressPercentage;
            progressBar.Update(); 
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            int statt =  Stat;
            
            button1.Enabled = true;
            button2.Enabled = true;
            if (statt != 0)
            {
                lblStatus.Text = "Upload complete!";
                MessageBox.Show("Upload Success", "Done",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            else
            {
                lblStatus.Text = "Invalid Credentials";
                MessageBox.Show("Invalid Credentials", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox2.Clear();
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
