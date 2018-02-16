using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;

namespace AutomateWebContentSelector
{
    public partial class Form1 : Form
    {
        DataGridView dgvCityDetails = new DataGridView();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if(dataGridView1.Rows.Count>1)
            {
                Automate();
            }
            else
            {
                MessageBox.Show("Make sure that the grid dont have any empty rows", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
          

        }

        private void Automate()
        {
            bool HasErrors = false;
            Guid g;
            g = Guid.NewGuid();
            string FileName = g.ToString();
          
            DataTable dt = new DataTable();
            dt.Columns.Add("Author name", typeof(string));
            dt.Columns.Add("University", typeof(string));
            dt.Columns.Add("H index", typeof(string));
            dt.Columns.Add("Document Name", typeof(string));


            try
            {
                IWebDriver driver = new ChromeDriver();
     
                foreach (DataGridViewRow row in this.dataGridView1.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value.ToString() != " ")
                        {

                            driver.Navigate().GoToUrl(cell.Value.ToString());
                            

                            string AuthorTextXPath = "//*[@id='authDetailsNameSection']/div/div[1]/div[1]/h2";
                            string UniversityTextXPath = "//*[@id='authDetailsNameSection']/div/div[1]/div[1]/div[1]";
                            string hindexTextXpath = "//*[@id='authorDetailsHindex']/div/div[2]/span";
                            string DocumentTextXpath = "//*[@id='authorDetailsDocumentsByAuthor']/div/div[2]/span";

                            /** Find the element **/
                            IWebElement h3Element = driver.FindElement(By.XPath(AuthorTextXPath));
                            IWebElement UniversityElement = driver.FindElement(By.XPath(UniversityTextXPath));
                            IWebElement hindexElement = driver.FindElement(By.XPath(hindexTextXpath));
                            IWebElement DocumentElement = driver.FindElement(By.XPath(DocumentTextXpath));


                            /** Grab the text **/
                            string AuthorName = h3Element.Text;
                            string Univerisity = UniversityElement.Text;
                            string Hindex = "h-index: " + hindexElement.Text;
                            string DocumentName = "Documents: " + DocumentElement.Text;
                            //driver.Close();
                            //driver.Dispose();
                            // PopulateRows(AuthorName, Univerisity, Hindex, DocumentName);
                            dt.Rows.Add(AuthorName, Univerisity, Hindex, DocumentName);
                        }


                    }
                }
                driver.Close();
                driver.Dispose();


            }
            catch (Exception ex)
            {

                //string strFilePath = string.Format(@"D:\Deeja_{0}.csv", FileName);
                //ToCSV(dt, strFilePath);
                //MessageBox.Show("File sucessfully generated and save to D drive");
                HasErrors = true;
                Task.Run(() =>
                {
                    var dialogResult = MessageBox.Show(ex.Message.ToString(), "Warning", MessageBoxButtons.OKCancel);
                    if (HasErrors)
                    {
                        MessageBox.Show("Some error occured Unable to complete the process" + Environment.NewLine + "Collected datas saved that was collected before the error occured", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }


                });

            }
            finally
            {
                string strFilePath = string.Format(@"D:\Deeja_{0}.csv", FileName);
                ToCSV(dt, strFilePath);
                if (!HasErrors)
                MessageBox.Show("File sucessfully generated and save to D drive");



            }
        }

        private void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
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
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
        private void PopulateRows(string AuthorName, string Univerisity, string Hindex, string DocumentName)
        {

            dgvCityDetails.ColumnCount = 4;
            dgvCityDetails.Columns[0].Name = "Product ID";
            dgvCityDetails.Columns[1].Name = "Product Name";
            dgvCityDetails.Columns[2].Name = "Product Price";

            string[] row = new string[] { "1", "Product 1", "1000" };
            dgvCityDetails.Rows.Add(row);
            row = new string[] { AuthorName, Univerisity, Hindex, DocumentName };
            dgvCityDetails.Rows.Add(row);

        }


        private void ExportToExcel()
        {
            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "ExportedFromDatGrid";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (int i = 0; i < dgvCityDetails.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dgvCityDetails.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check. 
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dgvCityDetails.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dgvCityDetails.Rows[i].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select file";
            fdlg.InitialDirectory = @"c:\";
            fdlg.FileName = textBox1.Text;
            fdlg.Filter = "Excel Sheet(*.xls)|*.xls|All Files(*.*)|*.*";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fdlg.FileName;
            }

            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.DataSet DtSet;
            System.Data.OleDb.OleDbDataAdapter MyCommand;
            MyConnection = new System.Data.OleDb.OleDbConnection(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + textBox1.Text + "';Extended Properties=Excel 8.0;");
            MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
            MyCommand.TableMappings.Add("Table", "Net-informations.com");
            DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);
            dataGridView1.DataSource = DtSet.Tables[0];
            MyConnection.Close();
        }

    }
}
