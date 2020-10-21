
using System;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using DGVPrinterHelper;


namespace DataGridView_Import_Excel
{
    public
    partial class Form1 : Form
    {
        private
            string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private
            string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";

        public
            Form1()
        {

            InitializeComponent();
            dataGridView1.Visible = false;
            btnPrint.Visible = false;
            button1.Visible = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode= DataGridViewTriState.True;
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private
            void BtnSelect_Click(object sender, EventArgs e)
        {

            Cursor = Cursors.WaitCursor;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                dataGridView1.Visible = true;
                btnPrint.Visible = true;
                button1.Visible = true;
               
            }
            Cursor = Cursors.Arrow;
        }

        private
            void OpenFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            string header = "YES";
            string conStr, sheetName;

            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }

            //Get the name of the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    //sheetName = "data";
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    //sheetName1 = dtExcelSchema.Rows[2]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }
           

            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {

                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        try
                        {
                            DataTable dt = new DataTable();

                            cmd.CommandText = "SELECT * From [" + sheetName + "]";
                            cmd.Connection = con;
                            con.Open();
                            oda.SelectCommand = cmd;
                            oda.Fill(dt);
                            con.Close();

                            Saloon Q3_326 = new Saloon("Audi Q3");
                            Saloon G11 = new Saloon("G11");
                            Saloon G3 = new Saloon("G3");
                            Saloon BMWvoga = new Saloon("BMWvoga");
                            Saloon BMWhiga = new Saloon("BMWhiga");
                            Saloon BR223 = new Saloon("BR223");
                            Saloon Skoda = new Saloon("SK38");



                            foreach (DataRow row in dt.Rows)
                            {
                                if (row[6].ToString().Contains("Audi") && row[6].ToString().Contains("Q3"))
                                {
                                    Calculation.Q3calc(row,ref Q3_326);
                                }
                                else if(row[6].ToString().ToUpper().Contains("G1"))
                                {
                                    Calculation.G11calc(row, ref G11);
                                }
                                else if(row[6].ToString().ToUpper().Contains("G3Y") || row[6].ToString().ToUpper().Contains("F90"))
                                {
                                    Calculation.G3calc(row, ref G3);
                                }
                                else if(row[6].ToString().ToUpper().Contains("BR223"))
                                {
                                    Calculation.BR223calc(row, ref BR223);
                                }
                                else if(row[6].ToString().ToUpper().Contains("SK38"))
                                {
                                    Calculation.SK38calc(row,ref Skoda);
                                }
                            }
                      
                            DataTable result = new DataTable();
                            result.Clear();
                            result.Columns.Add("Проект");
                            result.Columns.Add("Кількість чохлів").DataType = typeof(string);
                            result.Columns.Add("Загальний час").DataType=typeof(string);
                            result.Columns.Add("Час на одну штуку");
                            result.Columns.Add("Час на салон");
                            result.Columns.Add("Кількість салонів").DataType = typeof(int);
                            result.Columns.Add("Середній час на одну штуку");
                            result.Columns.Add("Коефіцієнт/кількість компонентів");
                            result.Columns.Add("Кількість компонент помножено на середній на одну штуку");
                            result.Columns.Add("Prod. sets planned").DataType = typeof(double);
                            //result.Columns.Add("Кількість комлектних салонів");
                            List<Saloon> Cars = new List<Saloon>();
                            
                            Q3_326.RCcount = Q3_326.RC40count + Q3_326.RC60count;
                            Q3_326.RCtime = Q3_326.RC40time + Q3_326.RC60time;
                            Q3_326.Coef = 9.0;
                            Cars.Add(Q3_326);
                            G11.Coef = 7.5;
                            G3.Coef = 4.0;
                            Cars.Add(G11);
                            Cars.Add(G3);

                            BMWvoga.Coef = 4.0;
                            BMWhiga.Coef = 3.5;
                            Calculation.BMWvogacalc(G11, G3,ref BMWvoga);
                            Calculation.BMWhigacalc(G11 ,ref BMWhiga);
                            
                            Cars.Add(BMWvoga);
                            Cars.Add(BMWhiga);

                            BR223.Coef = 8.0;
                            Cars.Add(BR223);

                            Skoda.Coef = 3.0;
                            Cars.Add(Skoda);



                            foreach (Saloon car in Cars)
                            {
                                DataRow row1 = result.NewRow();
                                if (car.ProjectName == "Audi Q3")
                                {
                                    FormatRow.Q3row(car, ref row1);
                                }
                                else if(car.ProjectName == "G11")
                                {
                                    FormatRow.G11row(car, ref row1);
                                }
                                else if(car.ProjectName == "G3")
                                {
                                    FormatRow.G3row(car, ref row1);
                                }
                                else if (car.ProjectName == "BR223")
                                {
                                    FormatRow.BR223row(car, ref row1);
                                }
                                else if(car.ProjectName == "SK38")
                                {
                                    FormatRow.SK38row(car, ref row1);
                                }
                                else if(car.ProjectName == "BMWhiga")
                                {
                                    FormatRow.BMWhiga(car, ref row1);
                                }
                                else if( car.ProjectName == "BMWvoga")
                                {
                                    FormatRow.BMWvoga(car, ref row1);
                                }
                                row1["Коефіцієнт/кількість компонентів"] =  car.Coef;
                                row1["Кількість компонент помножено на середній на одну штуку"] = Math.Round(car.Coef * car.AVGtime, 3);
                                row1["Prod. sets planned"] = Math.Round( 480/ (car.Coef * car.AVGtime),3);
                                //row1["Кількість комлектних салонів"] = car.CompleteSaloons();
                                result.Rows.Add(row1);
                            }

                           
                            dataGridView1.DataSource = result;
                            Cursor = Cursors.Arrow;
                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
                        }
                    }
                }
            }
        }

        private
            void Form1_Load(object sender, EventArgs e)
        {
        }

        private
            void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //foreach (DataGridViewRow Myrow in dataGridView1.Rows)
            //{ //Here 2 cell is target value and 1 cell is Volume
            //    if (Convert.ToInt32(Myrow.Cells[5].Value) < 0) // Or your condition
            //    {
            //        Myrow.Cells[5].Style.BackColor = Color.Red;
            //    }
            //    else
            //    {
            //        //Myrow.DefaultCellStyle.BackColor = Color.Green;
            //    }
            //}
        }
        private
            void Button1_Click(object sender, EventArgs e)

        {
            SaveFileDialog openDlg = new SaveFileDialog();
            openDlg.Filter = "Execl files (*.xls)|*.xls";

            string path = openDlg.FileName;
            if (openDlg.ShowDialog() == DialogResult.OK)
            {

                Excel.Application xlApp;

                Excel.Workbook xlWorkBook;

                Excel.Worksheet xlWorkSheet;

                object misValue = System.Reflection.Missing.Value;
            System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

                Int16 i, j;

                xlApp = new Excel.ApplicationClass();
                xlApp.DisplayAlerts = false;

                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //xlWorkSheet.Cells[1, 1] = "Проект";
                //xlWorkSheet.Cells[1, 2] = "Кількість чохлів";
                //xlWorkSheet.Cells[1, 3] = "Загальний час";
                //xlWorkSheet.Cells[1, 4] = "Час на одну штуку";
                //xlWorkSheet.Cells[1, 5] = "Час на cалон";

                for (i = 0; i <= dataGridView1.RowCount - 1; i++)

                {

                    for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)

                    {

                        xlWorkSheet.Cells[i + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                    }
                }

                try
                {
                    xlWorkBook.SaveAs(path.ToString(), Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Сталася помилка під час збереження " + ex.ToString());
                }
                xlWorkBook.Close(true, misValue, misValue);

                xlApp.Quit();

                ReleaseObject(xlWorkSheet);

                ReleaseObject(xlWorkBook);

                ReleaseObject(xlApp);
            }
        }

        private
            void ReleaseObject(object obj)

        {

            try

            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);

                obj = null;
            }

            catch (Exception ex)

            {

                obj = null;

                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }

            finally

            {

                GC.Collect();
            }
        }

        private
            void BtnPrint_Click_1(object sender, EventArgs e)
        {
            DGVPrinter printer = new DGVPrinter();

            printer.Title = "Продуктивність";
            //printer.SubTitle = string.Format("Дата {0}", DateTime.Now);

            printer.SubTitleFormatFlags = StringFormatFlags.LineLimit |

                StringFormatFlags.NoClip;

            printer.PageNumbers = true;

            printer.PageNumberInHeader = false;


            printer.HeaderCellAlignment = StringAlignment.Center;
            printer.ColumnWidths.Add("Проект",90);
            printer.ColumnWidths.Add("Кількість чохлів", 90);
            printer.ColumnWidths.Add("Загальний час", 90);
            printer.ColumnWidths.Add("Час на одну штуку", 110);
            printer.ColumnWidths.Add("Час на салон", 90);
            printer.ColumnWidths.Add("Кількість салонів", 90);
            printer.ColumnWidths.Add("Середній час на одну штуку", 90);
            printer.Footer = "BADER";

            printer.FooterSpacing = 15;
            //printer.ColumnWidth = DGVPrinter.ColumnWidthSetting.DataWidth;

            printer.PrintDataGridView(dataGridView1);
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }
    }
}


