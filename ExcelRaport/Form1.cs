using System;
using System.Diagnostics;
using System.Globalization;
using System.Configuration;
using System.Collections.Specialized;
using System.IO;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using CheckBox = System.Windows.Forms.CheckBox;

namespace ExcelRaport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

       

        public string filepath = Application.StartupPath;
        public string file1 { get; private set; }
        public string projectName { get; private set; }

        public void config()
        {
            string sAttr;

            // Read all the keys from the config file
            NameValueCollection sAll;
            sAll = ConfigurationManager.AppSettings;

            foreach (string s in sAll.AllKeys)
                MessageBox.Show("Key: " + s + " Value: " + sAll.Get(s), "Error",
             MessageBoxButtons.OK, MessageBoxIcon.Error);

        }

        public void load_files(string customPath, string file1)
        {

            try
            {
                string cell1 = ConfigurationManager.AppSettings.Get("cell1");
                string cell2 = ConfigurationManager.AppSettings.Get("cell2");
                string cell3 = ConfigurationManager.AppSettings.Get("cell3");
                string destination = ConfigurationManager.AppSettings.Get("destination");


                XLWorkbook workbook = new XLWorkbook(customPath + file1);//(filepath + "\\Data\\" + file1);
                XLWorkbook workbook2 = new XLWorkbook();
                IXLWorksheet worksheet2 = workbook2.AddWorksheet("Sheet1");

                Guid UUID = Guid.NewGuid();

                string filename2 = "wyjsciowe z____" + file1.Substring(0, file1.IndexOf(".")) + "____" + UUID + ".xlsx"; //+ DateTime.Now.ToString("-fff-ss-mm-HH-dd-MM-yyyy")
                int i = 2;

                //string cell1 = textBox1.Text.ToString();
                //string cell2 = textBox2.Text.ToString();
                //string cell3 = textBox3.Text.ToString();
                foreach (IXLWorksheet worksheet in workbook.Worksheets)
                {
                    try
                    {
                        var cell_date = worksheet.Cell(cell1).Value;
                        DateTime data_wydania = DateTime.Parse(cell_date.ToString());
                        Console.WriteLine(cell_date);
                        var name = worksheet.Cell(cell2).Value;
                        var number = worksheet.Cell(cell3).Value;
                        Console.WriteLine(cell_date.ToString() + " " + name.ToString() + " " + number);
                        worksheet2.Cell(1, 1).SetValue("Numer telefonu:");
                        worksheet2.Cell(1, 2).SetValue("Imię i nazwisko:");
                        worksheet2.Cell(1, 3).SetValue("Data wydania karty:");
                        worksheet2.Cell(1, 4).SetValue("Data wygaśnięcia:");
                        worksheet2.Cell(1, 5).SetValue("Pozostało dni:");
                        worksheet2.Cell(i, 3).SetValue(data_wydania);
                        worksheet2.Cell(i, 2).SetValue(name);
                        worksheet2.Cell(i, 1).SetValue(number);
                        //DateTime data_wygasniecia = DateTime.ParseExact(cell_date.ToString(), "yyyy-MM-dd", null);
                        //Console.WriteLine(Output);

                        DateTime data_wygasniecia = DateTime.Parse(cell_date.ToString()).AddYears(5);
                        double wygasniecie = (data_wygasniecia - DateTime.Today).TotalDays;
                        worksheet2.Cell(i, 4).SetValue(data_wygasniecia);
                        worksheet2.Cell(i, 5).SetValue(wygasniecie);

                        if (wygasniecie <= 7.0)
                        {
                            worksheet2.Cell(i, 5).Style.Fill.BackgroundColor = XLColor.CandyAppleRed;
                        }
                        else if (wygasniecie <= 30.0)
                        {
                            worksheet2.Cell(i, 5).Style.Fill.BackgroundColor = XLColor.ChromeYellow;
                        }
                        else if (wygasniecie <= 90.0)
                        {
                            worksheet2.Cell(i, 5).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                        else if (wygasniecie <= 180.0)
                        {
                            worksheet2.Cell(i, 5).Style.Fill.BackgroundColor = XLColor.AppleGreen;
                        }
                        else
                        {
                            worksheet2.Cell(i, 5).Style.Fill.BackgroundColor = XLColor.Green;
                        }



                        i++;
                    }
                    catch (Exception ee)
                    {
                        Console.WriteLine("Exception: " + ee.Message);

                    }

                }
                workbook.SaveAs(file1);
                worksheet2.Columns().AdjustToContents();
                worksheet2.Rows().AdjustToContents();
                workbook2.SaveAs(destination + filename2);
                FileInfo fi = new FileInfo(destination + filename2);
                if (fi.Exists)
                {
                    Process.Start(destination + filename2);
                }
            }

            catch (Exception ee)
            {
                Console.WriteLine("Exception: " + ee.Message);
                MessageBox.Show("Wystąpił błąd. Być może nie został wybrany żaden plik z listy.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        

        public void refresh_default()
        {
            tableLayoutPanel1.Controls.Clear();

            string customPath = ConfigurationManager.AppSettings.Get("customPath");

            try
            {
                string[] files = Directory.GetFiles(customPath, "*.xlsx");//filepath + "\\Data", );
                int i = 0;
                string[] array = files;
                foreach (string file in array)
                {
                    string fullPath = Path.GetFullPath(file).TrimEnd(Path.DirectorySeparatorChar);
                    projectName = Path.GetFileName(fullPath);

                    

                    CheckBox checkBox = new CheckBox
                    {
                        Text = projectName

                    };
                    checkBox.AutoSize = true;
                    checkBox.Click += label1_Click;
                    tableLayoutPanel1.Controls.Add(checkBox, 0, i);



                    i++;
                    i++;
                }


            }
            catch(Exception ee)
            {
                MessageBox.Show("Wystąpił błąd. Nie znaleziono ścieżki. Została ustawiona domyślna ścieżka.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);

                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings["customPath"].Value = filepath + "\\";
                config.AppSettings.Settings["destination"].Value = filepath + "\\";
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
                

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            refresh_default();  
        }
        private void label1_Click(object sender, EventArgs e)
        {
            //RadioButton locaLRadioButton = (RadioButton)sender;
            CheckBox localCheckBox = (CheckBox)sender;

            //file1 = locaLRadioButton.Text.ToString(); 
            file1 = localCheckBox.Text.ToString();
        }
        private void ustawieniaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.Owner = this;
         
            if (Application.OpenForms["Form2"] == null)
            {
                frm.Show();
            }           
        }


        private void wczytajToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string customPath = ConfigurationManager.AppSettings.Get("customPath");
            load_files(customPath, file1);
        }

        private void wczytajZWybranejLokalizacjiToolStripMenuItem_Click(object sender, EventArgs e)
        {
           

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void odśiweżToolStripMenuItem_Click(object sender, EventArgs e)
        {
            refresh_default();
        }

        private void wczytajZPodanejLokalizacjiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog folderBrowser = new OpenFileDialog();
            folderBrowser.CheckFileExists = true;
            folderBrowser.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                string customPath = Path.GetDirectoryName(folderBrowser.FileName) + "\\";
                string file1 = folderBrowser.SafeFileName;
                load_files(customPath, file1);
            }
        }

        private void wczytajLokalizacjeToolStripMenuItem_Click(object sender, EventArgs e)
        {
                       
        }

        private void testToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
    }
}
