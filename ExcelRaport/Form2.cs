using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelRaport
{
    
    public partial class Form2 : Form
    {
        public string filepath = Application.StartupPath;
        public Form2()
        {
            InitializeComponent();
        }

        

        private void Form2_Load(object sender, EventArgs e)
        {
            textBox1.Text = ConfigurationManager.AppSettings.Get("cell1");
            textBox2.Text = ConfigurationManager.AppSettings.Get("cell2");
            textBox3.Text = ConfigurationManager.AppSettings.Get("cell3");

            textBox4.Text = ConfigurationManager.AppSettings.Get("customPath");
            textBox5.Text = ConfigurationManager.AppSettings.Get("destination");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            //make changes
            config.AppSettings.Settings["cell1"].Value = textBox1.Text.ToString();
            config.AppSettings.Settings["cell2"].Value = textBox2.Text.ToString();
            config.AppSettings.Settings["cell3"].Value = textBox3.Text.ToString();

            config.AppSettings.Settings["customPath"].Value = textBox4.Text.ToString();
            config.AppSettings.Settings["destination"].Value = textBox5.Text.ToString();


            //save to apply changes
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

           
            Form1 frm1 = (Form1)this.Owner;
            frm1.refresh_default();

            this.Close();
            
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {


            OpenFileDialog folderBrowser = new OpenFileDialog();
           
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;

            folderBrowser.FileName = "Wybór katalogu";


            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                string folderPath = Path.GetDirectoryName(folderBrowser.FileName);
                textBox4.Text = folderPath +"\\";
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            


                OpenFileDialog folderBrowser = new OpenFileDialog();

                folderBrowser.ValidateNames = false;
                folderBrowser.CheckFileExists = false;
                folderBrowser.CheckPathExists = true;

                folderBrowser.FileName = "Wybór katalogu";


                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = Path.GetDirectoryName(folderBrowser.FileName);
                    textBox5.Text = folderPath + "\\";
                }
            
        }
    }
}
