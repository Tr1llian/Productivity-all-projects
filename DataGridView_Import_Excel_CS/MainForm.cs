
using System;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Drawing;
using DGVPrinterHelper;
using ClosedXML.Excel;
using System.Text;
using Productivity;
using System.ComponentModel;

namespace DataGridView_Import_Excel
{
    public partial class Form1 : Form
    {
        List<Saloon> Cars = new List<Saloon>();
    
        private readonly string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private readonly string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
       

        public
            Form1()
        {

            InitializeComponent();


            dataGridView1.Visible = false;
            btnPrint.Visible = false;
            button1.Visible = false;
            button2.Visible = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.DefaultCellStyle.WrapMode= DataGridViewTriState.True;
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
          

        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {

            Cursor = Cursors.WaitCursor;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
            }
            Cursor = Cursors.Arrow;
        }

        private void OpenFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
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
                    try
                    {
                        con.Open();
                    }
                    catch(Exception ex)
                    {

                        Console.WriteLine(ex.ToString());
                        MyMessages m = new MyMessages("Невірний формат", "Упевніться, що ексель має розширення xlsx і має структуру як на фото нижче", 2);
                        m.ShowDialog();

                        bool okButtonClicked = m.OKButtonClicked;
                        return;
                    }
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    //sheetName = "data";
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    //sheetName1 = dtExcelSchema.Rows[2]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }

            dataGridView1.Visible = true;
            btnPrint.Visible = true;
            button1.Visible = true;
            button2.Visible = true;
            btnSelect.Visible = false;

            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {

                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        try
                        {

                            WaitForm wf = new WaitForm();
                            wf.Show();
                            DataTable dt = new DataTable();
                            cmd.CommandText = "SELECT * From [" + sheetName + "]";
                            cmd.Connection = con;
                            con.Open();
                            oda.SelectCommand = cmd;
                            oda.Fill(dt);
                            con.Close();

                            Saloon Q3 = new SaloonQ3("Audi Q3");
                            Saloon G11 = new SaloonG11("G11");
                            Saloon G3 = new SaloonG3("G3");
                            Saloon BR223 = new SaloonBR223("BR223");
                            Saloon Skoda = new SaloonSK38("SK38");

                            int i = 0;
                            int percentcoef = dt.Rows.Count / 100;
                            foreach (DataRow row in dt.Rows)
                            {
                                i++;
                                if (row[6].ToString().Contains("Audi") && row[6].ToString().Contains("Q3"))
                                {
                                    Q3.ParseExcel(row);
                                }
                                else if (row[6].ToString().ToUpper().Contains("G1"))
                                {
                                    G11.ParseExcel(row);
                                }
                                else if (row[6].ToString().ToUpper().Contains("G3Y") || row[6].ToString().ToUpper().Contains("F90"))
                                {
                                    G3.ParseExcel(row);
                                }
                                else if (row[6].ToString().ToUpper().Contains("BR223"))
                                {
                                    BR223.ParseExcel(row);
                                }
                                else if (row[6].ToString().ToUpper().Contains("SK38"))
                                {
                                    Skoda.ParseExcel(row);
                                }

                                if (wf.ProgressBarValue ==(int) (i / percentcoef))
                                {
                                    continue;
                                }
                                else
                                {
                                    wf.ProgressBarValue++;
                                    Console.WriteLine(wf.ProgressBarValue.ToString());
                                }

                            }

                            Saloon BMWhiga = new SaloonBMWhiga(G11, "BMWhiga");
                            Saloon BMWvoga = new SaloonBMWvoga(G11, G3, "BMWvoga");

                            foreach (DataRow row in dt.Rows)
                            {
                                if (row[6].ToString().ToUpper().Contains("G1") || row[6].ToString().ToUpper().Contains("G3Y") || row[6].ToString().ToUpper().Contains("F90"))
                                {
                                    if (row[6].ToString().ToUpper().Contains("RC") || row[6].ToString().ToUpper().Contains("RB"))
                                    {
                                        BMWhiga.ParseExcel(row);
                                    }
                                    else if (row[6].ToString().ToUpper().Contains("FC") || row[6].ToString().ToUpper().Contains("FB"))
                                    {
                                        BMWvoga.ParseExcel(row);
                                    }
                                }
                            }

                            DataTable result = new DataTable();
                            result.Clear();
                            result.Columns.Add("Проект");
                            result.Columns.Add("Кількість чохлів").DataType = typeof(string);
                            result.Columns.Add("Загальний час").DataType = typeof(string);
                            result.Columns.Add("Час на одну штуку");
                            result.Columns.Add("Час на салон");
                            result.Columns.Add("Кількість салонів").DataType = typeof(int);
                            result.Columns.Add("Середній час на одну штуку").DataType = typeof(double);
                            result.Columns.Add("Коефіцієнт/кількість компонентів").DataType = typeof(double);
                            result.Columns.Add("Кількість компонент помножено на середній на одну штуку").DataType = typeof(double);
                            result.Columns.Add("Prod. sets planned").DataType = typeof(double);
                            result.Columns.Add("Кількість бригад").DataType = typeof(int);
                            result.Columns.Add("Кількість днів").DataType = typeof(int);
                            result.Columns.Add("Кількість бригад soll").DataType = typeof(int);
                            result.Columns.Add("Кількість бригад ist").DataType = typeof(int);
                            result.Columns.Add("Коефіцієнт").DataType = typeof(double);
                            result.Columns.Add("Дні").DataType = typeof(double);


                            Q3.Coef = 9.0;
                            Cars.Add(Q3);

                            G11.Coef = 8;
                            G3.Coef = 4.0;
                            Cars.Add(G11);
                            Cars.Add(G3);

                            BMWvoga.Coef = 4.0;
                            BMWhiga.Coef = 4.0;


                            Cars.Add(BMWvoga);
                            Cars.Add(BMWhiga);

                            BR223.Coef = 8.0;
                            Cars.Add(BR223);

                            Skoda.Coef = 3.0;
                            Cars.Add(Skoda);

                            LinesDayCount LD = new LinesDayCount(ref Cars);
                            LD.ShowDialog();

                            int SaloonSum = 0;
                            int LinesSum = 0;
                            int PlanLinesSum = 0;
                            double MiddleSetplan = 0;


                            foreach (Saloon car in Cars)
                            {
                                DataRow row1 = result.NewRow();
                                car.CreateRow(ref row1);
                                result.Rows.Add(row1);
                            }

                            foreach (DataRow a in result.Rows)
                            {
                                SaloonSum += Convert.ToInt32(a[5].ToString());
                                LinesSum += Convert.ToInt32(a[13].ToString());
                                PlanLinesSum += Convert.ToInt32(a[10].ToString());
                                MiddleSetplan += Convert.ToDouble(a[9].ToString()) * Convert.ToDouble(a[13].ToString());
                            }
                            DataRow row2 = result.NewRow();
                            row2[0] = "Підсумок";
                            row2[9] = Math.Round(MiddleSetplan / LinesSum, 3);
                            row2[5] = SaloonSum;
                            row2[13] = LinesSum;
                            row2[10] = PlanLinesSum;
                            result.Rows.Add(row2);
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

    
        private void Form1_Load(object sender, EventArgs e)
        {
        }


        public  void CreateExcel(bool a)
        {
            var workbook = new XLWorkbook();
            workbook.AddWorksheet("Wochenbericht");
            var ws = workbook.Worksheet("Wochenbericht");

            int row = 1;
            ws.Cell("A" + row.ToString()).Value = "Проект";
            ws.Cell("B" + row.ToString()).Value = "Кількість чохлів";
            ws.Cell("C" + row.ToString()).Value = "Загальний час";
            ws.Cell("D" + row.ToString()).Value = "Час на одну штуку";
            ws.Cell("E" + row.ToString()).Value = "Час на салон";
            ws.Cell("F" + row.ToString()).Value = "Кількість салонів";
            ws.Cell("G" + row.ToString()).Value = "Середній час на одну штуку";
            ws.Cell("H" + row.ToString()).Value = "Коефіцієнт/кількість компонентів";
            StringBuilder str = new StringBuilder();
            str.Append("Кількість компонент помножено");
            str.AppendLine();
            str.Append("на середній на одну штуку");
            ws.Cell("I" + row.ToString()).Value = str.ToString();
            ws.Cell("J" + row.ToString()).Value = "Prod. sets planned";
            ws.Cell("K" + row.ToString()).Value = "Кількість бригад";
            ws.Cell("L" + row.ToString()).Value = "Кількість днів";
            ws.Cell("M" + row.ToString()).Value = "Кількість бригад soll";
            ws.Cell("N" + row.ToString()).Value = "Кількість бригад ist";
            ws.Cell("O" + row.ToString()).Value = "Коефіцієнт";
            ws.Cell("P" + row.ToString()).Value = "дні";

            var rngTable = ws.Range("A1:P1");
            rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngTable.Style.Font.Bold = true;
            rngTable.Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);
            row = 2;
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                if (item.Cells[0].Value.ToString() == "Підсумок")
                {
                    continue;
                }
                else
                {
                    ws.Cell("A" + row.ToString()).Value = item.Cells[0].Value.ToString();
                    ws.Cell("B" + row.ToString()).Value = item.Cells[1].Value.ToString();
                    ws.Cell("C" + row.ToString()).Value = item.Cells[2].Value.ToString();
                    ws.Cell("D" + row.ToString()).Value = item.Cells[3].Value.ToString();
                    ws.Cell("E" + row.ToString()).Value = item.Cells[4].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("F" + row.ToString()).Value = item.Cells[5].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("G" + row.ToString()).Value = item.Cells[6].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("H" + row.ToString()).Value = item.Cells[7].Value.ToString().Replace(',', '.');
                    ws.Cell("I" + row.ToString()).Value = item.Cells[8].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("J" + row.ToString()).Value = item.Cells[9].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("K" + row.ToString()).Value = item.Cells[10].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("L" + row.ToString()).Value = item.Cells[11].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("M" + row.ToString()).Value = item.Cells[12].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("N" + row.ToString()).Value = item.Cells[13].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("O" + row.ToString()).Value = item.Cells[14].Value.ToString().Replace(',', '.'); ;
                    ws.Cell("P" + row.ToString()).Value = item.Cells[15].Value.ToString().Replace(',', '.'); ;
                    row++;
                }
            }
            ws.Cell("A" + row.ToString()).Value = "Підсумок";
            ws.Cell("F" + row.ToString()).FormulaA1 = "=SUM(F2:F8)";
            ws.Cell("F" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 255, 0);
            ws.Cell("N" + row.ToString()).FormulaA1 = "=SUM(N2:N8)";
            ws.Cell("N" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 255, 0);
            ws.Cell("K" + row.ToString()).FormulaA1 = "=SUM(K2:K8)";
            ws.Cell("K" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 255, 0);
            ws.Cell("P" + row.ToString()).FormulaA1 = "=(P2*K2+P5*K5+P6*K6+P7*K7+P8*K8)/K9";
            ws.Cell("P" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 255, 0);
            ws.Cell("J" + row.ToString()).FormulaA1 = "=(N2*J2+J5*N5+J6*N6+J7*N7+J8*N8)/N9";
            ws.Cell("J" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 255, 0);
            ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            ws.Columns().AdjustToContents();
            
               
            ws.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            IXLRange titleRange = ws.Range("A1:P9");

            titleRange.Cells().Style
                .Alignment.SetWrapText(true); // Its single statement
            titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Columns().Width = 11;
            ws.Rows().Height = 30;
            ws.Row(1).Height = 83;
            titleRange.Cells().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;


            foreach(Saloon car in Cars)
            {
                if(car.ProjectName == "G11" || car.ProjectName == "G3")
                {
                    continue;
                }
                workbook.AddWorksheet("Бригади "+car.ProjectName);
                ws = workbook.Worksheet("Бригади " + car.ProjectName);
                int row2 = 1;
                ws.Cell("A" + row2.ToString()).Value = "День";
                ws.Cell("B" + row2.ToString()).Value = "Бригада";
                ws.Cell("A" + row2.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);
                ws.Cell("B" + row2.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);
                row2 = 2;
                car.LD.Sort();
                foreach(LineDay ld in car.LD)
                {
                    ws.Cell("A" + row2.ToString()).Value = ld.Date;
                    ws.Cell("B" + row2.ToString()).Value = ld.Name;
                    row2++;
                }
                ws.Columns().Width = 15;
                ws.Cells().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cells().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            }

            if (a == false)
            {
                workbook.SaveAs(@"C:/test/productivity.xlsx");
            }
            else
            {
                saveFileDialog1.Filter = "*.xlsx|";
                _ = saveFileDialog1.FileName;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialog1.FileName + ".xlsx");
                }
            }
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            CreateExcel(true);
           
        }

        private void BtnPrint_Click_1(object sender, EventArgs e)
        {
            DGVPrinter printer = new DGVPrinter
            {
                Title = "Продуктивність",
                //printer.SubTitle = string.Format("Дата {0}", DateTime.Now);

                SubTitleFormatFlags = StringFormatFlags.LineLimit |

                StringFormatFlags.NoClip,

                PageNumbers = true,

                PageNumberInHeader = false,

                HeaderCellAlignment = StringAlignment.Center
            };

            printer.ColumnWidths.Add("Проект",70);
            printer.ColumnWidths.Add("Кількість чохлів", 90);
            printer.ColumnWidths.Add("Загальний час", 100);
            printer.ColumnWidths.Add("Час на одну штуку", 110);
            printer.ColumnWidths.Add("Час на салон", 60);
            printer.ColumnWidths.Add("Кількість салонів", 60);
            printer.ColumnWidths.Add("Середній час на одну штуку", 40);
            printer.ColumnWidths.Add("Коефіцієнт/кількість компонентів", 40);
            printer.ColumnWidths.Add("Кількість компонент помножено на середній на одну штуку", 70);
            printer.ColumnWidths.Add("Prod. sets planned", 40);
            printer.Footer = "BADER";

            printer.FooterSpacing = 15;
            //printer.ColumnWidth = DGVPrinter.ColumnWidthSetting.DataWidth;

            printer.PrintDataGridView(dataGridView1);
        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            CreateExcel(false);
            var LoginGW = new LoginGW();
            _ = LoginGW.ShowDialog();
            
        }
    }
}


