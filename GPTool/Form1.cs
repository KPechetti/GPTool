using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GPTool
{
    public partial class GPTool : Form
    {
        public GPTool()
        {
           
            InitializeComponent();
            ErrorMsg.Text = "";
            printDocument1.DefaultPageSettings.Landscape = true;
        }

        private string billNo, year, assesNum, jilla, doorNo, mandal, name, amountInText, panchayat;
        private string houseTax, waterTax, libSes, lightTax, drainageTax, totalTax;

        private void billNoFrom_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
        (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void billNoTo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
        (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void billNum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
        (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ErrorMsg.Text = "Processing...";
            string path = excelPath.Text;
            billNo = billNum.Text;
            if(string.IsNullOrEmpty(billNo))
            {
                MessageBox.Show("Please enter Assesment Number");
                return;
            }
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("Please enter Excel Path");
                return;
            }
            string  status =  ReadExcel(path, assesNum, "PrintOne");
            printDocument1.DefaultPageSettings.Landscape = true;
                //if (printPreviewDialog1.ShowDialog() == DialogResult.OK)
            if(status == "completed")
            {
                ErrorMsg.Text = "Printing...";
                printDocument1.Print();
                ErrorMsg.Text = "Completed.";
            }
            
        }

        private void printAll_Click(object sender, EventArgs e)
        {
            ErrorMsg.Text = "Processing...";
            string path = excelPath.Text;
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("Please enter Excel Path");
                return;
            }            
            string status = ReadExcel(path, assesNum, "PrintAll");
            ErrorMsg.Text = "Completed.";
        }

        private void printRange_Click(object sender, EventArgs e)
        {
            int billFrom = int.Parse(billNoFrom.Text);
            int billTo = int.Parse(billNoTo.Text);
            ErrorMsg.Text = "Processing...";
            string path = excelPath.Text;

            if (string.IsNullOrEmpty(billNoFrom.Text))
            {
                MessageBox.Show("Please enter bill Number from");
                return;
            }
            if (string.IsNullOrEmpty(billNoTo.Text))
            {
                MessageBox.Show("Please enter bill Number to");
                return;
            }
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("Please enter Excel Path");
                return;
            }
            string status = ReadExcel(path, assesNum, "PrintRange");
            //printDocument1.DefaultPageSettings.Landscape = true;
            ErrorMsg.Text = "Completed.";
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            System.Drawing.Font arialFont_panchayat = new System.Drawing.Font("Arial", 15, System.Drawing.FontStyle.Bold);
            System.Drawing.Image img = Properties.Resources.GPImage;
            Point loc = new Point(0, 0);
            e.Graphics.DrawImage(img, loc);

            using (System.Drawing.Font arialFont = new System.Drawing.Font("Arial", 7, System.Drawing.FontStyle.Bold))
            {
                StringFormat format = new StringFormat(StringFormatFlags.DirectionRightToLeft);
                //First copy
                PointF panchayat_loc1 = new PointF(202f, 37f);
                PointF billNo_loc1 = new PointF(157f, 70f);
                PointF assesNo_loc1 = new PointF(157f, 112f);
                PointF doorNo_loc1 = new PointF(157f, 148f);
                PointF year_loc1 = new PointF(368f, 74f);
                PointF jilla_loc1 = new PointF(368f, 113f);
                PointF mandal_loc1 = new PointF(368f, 153f);
                PointF name_loc1 = new PointF(43f, 243f);
                PointF totalTax_loc1 = new PointF(162f, 305f);
                PointF houseTax_loc1 = new PointF(444f, 412f);
                PointF waterTax_loc1 = new PointF(444f, 436f);
                PointF libSes_loc1 = new PointF(444f, 462f);
                PointF lightTax_loc1 = new PointF(444f, 496f);
                PointF drainageTax_loc1 = new PointF(444f, 525f);
                PointF totalTax_loc1_a = new PointF(444f, 604f);
                PointF amountInText_loc1 = new PointF(89f, 346f);

                e.Graphics.DrawString(panchayat, arialFont_panchayat, Brushes.Black, panchayat_loc1);
                e.Graphics.DrawString(billNo, arialFont, Brushes.Black, billNo_loc1);
                e.Graphics.DrawString(assesNum, arialFont, Brushes.Black, assesNo_loc1);
                e.Graphics.DrawString(doorNo, arialFont, Brushes.Black, doorNo_loc1);
                e.Graphics.DrawString(year, arialFont, Brushes.Black, year_loc1);
                e.Graphics.DrawString(jilla, arialFont, Brushes.Black, jilla_loc1);
                e.Graphics.DrawString(mandal, arialFont, Brushes.Black, mandal_loc1);
                e.Graphics.DrawString(name, arialFont, Brushes.Black, name_loc1);
                e.Graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc1);
                e.Graphics.DrawString(houseTax, arialFont, Brushes.Black, houseTax_loc1, format);
                e.Graphics.DrawString(waterTax, arialFont, Brushes.Black, waterTax_loc1, format);
                e.Graphics.DrawString(libSes, arialFont, Brushes.Black, libSes_loc1, format);
                e.Graphics.DrawString(lightTax, arialFont, Brushes.Black, lightTax_loc1, format);
                e.Graphics.DrawString(drainageTax, arialFont, Brushes.Black, drainageTax_loc1, format);
                e.Graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc1_a, format);
                e.Graphics.DrawString(amountInText, arialFont, Brushes.Black, amountInText_loc1);

                //Second copy
                PointF panchayat_loc2 = new PointF(704f, 46f);
                PointF billNo_loc2 = new PointF(663f, 85f);
                PointF assesNo_loc2 = new PointF(663f, 122f);
                PointF doorNo_loc2 = new PointF(663f, 160f);
                PointF year_loc2 = new PointF(864f, 82f);
                PointF jilla_loc2 = new PointF(864f, 118f);
                PointF mandal_loc2 = new PointF(864f, 160f);
                PointF name_loc2 = new PointF(558f, 248f);
                PointF totalTax_loc2 = new PointF(671f, 308f);
                PointF houseTax_loc2 = new PointF(930f, 413f);
                PointF waterTax_loc2 = new PointF(930f, 433f);
                PointF libSes_loc2 = new PointF(930f, 466f);
                PointF lightTax_loc2 = new PointF(930f, 502f);
                PointF drainageTax_loc2 = new PointF(930f, 529f);
                PointF totalTax_loc2_a = new PointF(930f, 601f);
                PointF amountInText_loc2 = new PointF(597f, 350f);

                e.Graphics.DrawString(panchayat, arialFont_panchayat, Brushes.Black, panchayat_loc2);
                e.Graphics.DrawString(billNo, arialFont, Brushes.Black, billNo_loc2);
                e.Graphics.DrawString(assesNum, arialFont, Brushes.Black, assesNo_loc2);
                e.Graphics.DrawString(doorNo, arialFont, Brushes.Black, doorNo_loc2);
                e.Graphics.DrawString(year, arialFont, Brushes.Black, year_loc2);
                e.Graphics.DrawString(jilla, arialFont, Brushes.Black, jilla_loc2);
                e.Graphics.DrawString(mandal, arialFont, Brushes.Black, mandal_loc2);
                e.Graphics.DrawString(name, arialFont, Brushes.Black, name_loc2);
                e.Graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc2);
                e.Graphics.DrawString(houseTax, arialFont, Brushes.Black, houseTax_loc2, format);
                e.Graphics.DrawString(waterTax, arialFont, Brushes.Black, waterTax_loc2, format);
                e.Graphics.DrawString(libSes, arialFont, Brushes.Black, libSes_loc2, format);
                e.Graphics.DrawString(lightTax, arialFont, Brushes.Black, lightTax_loc2, format);
                e.Graphics.DrawString(drainageTax, arialFont, Brushes.Black, drainageTax_loc2, format);
                e.Graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc2_a, format);
                e.Graphics.DrawString(amountInText, arialFont, Brushes.Black, amountInText_loc2);

                //Third copy
                PointF billNo_loc3 = new PointF(1167f, 69f);
                PointF panchayat_loc3 = new PointF(1207f, 40f);
                PointF assesNo_loc3 = new PointF(1167f, 110f);
                PointF doorNo_loc3 = new PointF(1167f, 150f);
                PointF year_loc3 = new PointF(1370f, 70f);
                PointF jilla_loc3 = new PointF(1370f, 107f);
                PointF mandal_loc3 = new PointF(1370f, 149f);
                PointF name_loc3 = new PointF(1053f, 233f);
                PointF totalTax_loc3 = new PointF(1163f, 302f);
                PointF houseTax_loc3 = new PointF(1440f, 404f);
                PointF waterTax_loc3 = new PointF(1440f, 423f);
                PointF libSes_loc3 = new PointF(1440f, 448f);
                PointF lightTax_loc3 = new PointF(1440f, 475f);
                PointF drainageTax_loc3 = new PointF(1440f, 504f);
                PointF totalTax_loc3_a = new PointF(1440f, 558f);
                PointF amountInText_loc3 = new PointF(1096f, 343f);

                e.Graphics.DrawString(panchayat, arialFont_panchayat, Brushes.Black, panchayat_loc3);
                e.Graphics.DrawString(billNo, arialFont, Brushes.Black, billNo_loc3);
                e.Graphics.DrawString(assesNum, arialFont, Brushes.Black, assesNo_loc3);
                e.Graphics.DrawString(doorNo, arialFont, Brushes.Black, doorNo_loc3);
                e.Graphics.DrawString(year, arialFont, Brushes.Black, year_loc3);
                e.Graphics.DrawString(jilla, arialFont, Brushes.Black, jilla_loc3);
                e.Graphics.DrawString(mandal, arialFont, Brushes.Black, mandal_loc3);
                e.Graphics.DrawString(name, arialFont, Brushes.Black, name_loc3);
                e.Graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc3);
                e.Graphics.DrawString(houseTax, arialFont, Brushes.Black, houseTax_loc3, format);
                e.Graphics.DrawString(waterTax, arialFont, Brushes.Black, waterTax_loc3, format);
                e.Graphics.DrawString(libSes, arialFont, Brushes.Black, libSes_loc3, format);
                e.Graphics.DrawString(lightTax, arialFont, Brushes.Black, lightTax_loc3, format);
                e.Graphics.DrawString(drainageTax, arialFont, Brushes.Black, drainageTax_loc3, format);
                e.Graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc3_a, format);
                e.Graphics.DrawString(amountInText, arialFont, Brushes.Black, amountInText_loc3);

            }
        }

        #region Read Excel
        /// <summary>
        /// Read excel
        /// </summary>
        /// <param name="path"></param>
        /// <param name="assesNum"></param>
        private string ReadExcel(string path, string assesNum, string type)
        {
            Excel.Application oExcel = null;
            Excel.Workbook WB = null;
            try
            {
                //create a instance for the Excel object  
                oExcel = new Excel.Application();
                //pass that to workbook object  
                WB = oExcel.Workbooks.Open(path);
                Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[1];
                //statement get the first cell value  
                var firstcellvalue = ((Excel.Range)wks.Cells[1, 1]).Value;
                Excel.Range xlRange = wks.UsedRange;
                Boolean found = false;
                
                for (int i = 2; i <= xlRange.Rows.Count; i++)
                {
                    int currBillNo = int.Parse(((Excel.Range)wks.Cells[i, 1]).Value.ToString());
                    if ( type == "PrintOne" && currBillNo.ToString() == billNo)
                    {
                            found = true;                     
                            assesNum = ((Excel.Range)wks.Cells[i, 2]).Value.ToString();
                            doorNo = ((Excel.Range)wks.Cells[i, 3]).Value.ToString();
                            name = ((Excel.Range)wks.Cells[i, 4]).Value.ToString();
                            houseTax = string.Format("{0:F2}",((Excel.Range)wks.Cells[i, 6]).Value);
                            libSes = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 7]).Value);
                            waterTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 8]).Value);
                            drainageTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 9]).Value);
                            lightTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 10]).Value);
                            totalTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 11]).Value);
                            year = ((Excel.Range)wks.Cells[i, 12]).Value.ToString();
                            jilla = ((Excel.Range)wks.Cells[i, 13]).Value.ToString();
                            mandal = ((Excel.Range)wks.Cells[i, 14]).Value.ToString();
                            panchayat = ((Excel.Range)wks.Cells[i, 15]).Value.ToString();
                            amountInText = ((Excel.Range)wks.Cells[i, 16]).Value.ToString();                        
                            break;
                    }
                    else if(type == "PrintAll")
                    {
                        assesNum = ((Excel.Range)wks.Cells[i, 2]).Value.ToString();
                        doorNo = ((Excel.Range)wks.Cells[i, 3]).Value.ToString();
                        name = ((Excel.Range)wks.Cells[i, 4]).Value.ToString();
                        houseTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 6]).Value);
                        libSes = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 7]).Value);
                        waterTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 8]).Value);
                        drainageTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 9]).Value);
                        lightTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 10]).Value);
                        totalTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 11]).Value);
                        year = ((Excel.Range)wks.Cells[i, 12]).Value.ToString();
                        jilla = ((Excel.Range)wks.Cells[i, 13]).Value.ToString();
                        mandal = ((Excel.Range)wks.Cells[i, 14]).Value.ToString();
                        panchayat = ((Excel.Range)wks.Cells[i, 15]).Value.ToString();
                        amountInText = ((Excel.Range)wks.Cells[i, 16]).Value.ToString();
                        //Print
                        printDocument1.Print();

                    }
                    else if(type == "PrintRange"  && currBillNo >= int.Parse(billNoFrom.Text) && currBillNo <= int.Parse(billNoTo.Text))
                    {
                        assesNum = ((Excel.Range)wks.Cells[i, 2]).Value.ToString();
                        doorNo = ((Excel.Range)wks.Cells[i, 3]).Value.ToString();
                        name = ((Excel.Range)wks.Cells[i, 4]).Value.ToString();
                        houseTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 6]).Value);
                        libSes = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 7]).Value);
                        waterTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 8]).Value);
                        drainageTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 9]).Value);
                        lightTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 10]).Value);
                        totalTax = string.Format("{0:F2}", ((Excel.Range)wks.Cells[i, 11]).Value);
                        year = ((Excel.Range)wks.Cells[i, 12]).Value.ToString();
                        jilla = ((Excel.Range)wks.Cells[i, 13]).Value.ToString();
                        mandal = ((Excel.Range)wks.Cells[i, 14]).Value.ToString();
                        panchayat = ((Excel.Range)wks.Cells[i, 15]).Value.ToString();
                        amountInText = ((Excel.Range)wks.Cells[i, 16]).Value.ToString();
                        //Print
                        printDocument1.Print();
                    }
                }
                if (!found) ErrorMsg.Text = "Bill Number Not found, Enter a valid number.";
                WB.Close(0);
                oExcel.Quit();
                oExcel = null;
                return "completed";
            }
            catch (Exception ex)
            {
                if (WB != null) WB.Close(0);
                if (oExcel != null) oExcel.Quit();
                ErrorMsg.Text = "There is an error.";
                MessageBox.Show(ex.Message, "Error");
                return "failed";

            }
        }
        #endregion

        #region editImage
        private void EditImage(string originalFile)
        {
            string imageFilePath = originalFile;
            Bitmap bitmap = (Bitmap)System.Drawing.Image.FromFile(imageFilePath);//load the image file
            StringFormat format = new StringFormat(StringFormatFlags.DirectionRightToLeft);

            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                using (System.Drawing.Font arialFont = new System.Drawing.Font("Arial", 7, System.Drawing.FontStyle.Bold))
                {
                    //First copy
                    PointF panchayat_loc1 = new PointF(202f, 37f);
                    PointF billNo_loc1 = new PointF(157f, 70f);
                    PointF assesNo_loc1 = new PointF(157f, 112f);
                    PointF doorNo_loc1 = new PointF(157f, 148f);
                    PointF year_loc1 = new PointF(368f, 74f);
                    PointF jilla_loc1 = new PointF(368f, 113f);
                    PointF mandal_loc1 = new PointF(368f, 153f);
                    PointF name_loc1 = new PointF(43f, 243f);
                    PointF totalTax_loc1 = new PointF(162f, 305f);
                    PointF houseTax_loc1 = new PointF(444f, 412f);
                    PointF waterTax_loc1 = new PointF(444f, 436f);
                    PointF libSes_loc1 = new PointF(444f, 462f);
                    PointF lightTax_loc1 = new PointF(444f, 496f);
                    PointF drainageTax_loc1 = new PointF(444f, 525f);
                    PointF totalTax_loc1_a = new PointF(444f, 604f);
                    PointF amountInText_loc1 = new PointF(89f, 346f);

                    graphics.DrawString(panchayat, arialFont, Brushes.Black, panchayat_loc1);
                    graphics.DrawString(billNo, arialFont, Brushes.Black, billNo_loc1);
                    graphics.DrawString(assesNum, arialFont, Brushes.Black, assesNo_loc1);
                    graphics.DrawString(doorNo, arialFont, Brushes.Black, doorNo_loc1);
                    graphics.DrawString(year, arialFont, Brushes.Black, year_loc1);
                    graphics.DrawString(jilla, arialFont, Brushes.Black, jilla_loc1);
                    graphics.DrawString(mandal, arialFont, Brushes.Black, mandal_loc1);
                    graphics.DrawString(name, arialFont, Brushes.Black, name_loc1);
                    graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc1);
                    graphics.DrawString(houseTax, arialFont, Brushes.Black, houseTax_loc1, format);
                    graphics.DrawString(waterTax, arialFont, Brushes.Black, waterTax_loc1, format);
                    graphics.DrawString(libSes, arialFont, Brushes.Black, libSes_loc1, format);
                    graphics.DrawString(lightTax, arialFont, Brushes.Black, lightTax_loc1, format);
                    graphics.DrawString(drainageTax, arialFont, Brushes.Black, drainageTax_loc1, format);
                    graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc1_a, format);
                    graphics.DrawString(amountInText, arialFont, Brushes.Black, amountInText_loc1);

                    //Second copy
                    PointF panchayat_loc2 = new PointF(704f, 46f);
                    PointF billNo_loc2 = new PointF(663f, 85f);
                    PointF assesNo_loc2 = new PointF(663f, 122f);
                    PointF doorNo_loc2 = new PointF(663f, 160f);
                    PointF year_loc2 = new PointF(864f, 82f);
                    PointF jilla_loc2 = new PointF(864f, 118f);
                    PointF mandal_loc2 = new PointF(864f, 160f);
                    PointF name_loc2 = new PointF(558f, 248f);
                    PointF totalTax_loc2 = new PointF(671f, 308f);
                    PointF houseTax_loc2 = new PointF(930f, 413f);
                    PointF waterTax_loc2 = new PointF(930f, 433f);
                    PointF libSes_loc2 = new PointF(930f, 466f);
                    PointF lightTax_loc2 = new PointF(930f, 502f);
                    PointF drainageTax_loc2 = new PointF(930f, 529f);
                    PointF totalTax_loc2_a = new PointF(930f, 601f);
                    PointF amountInText_loc2 = new PointF(597f, 350f);

                    graphics.DrawString(panchayat, arialFont, Brushes.Black, panchayat_loc2);
                    graphics.DrawString(billNo, arialFont, Brushes.Black, billNo_loc2);
                    graphics.DrawString(assesNum, arialFont, Brushes.Black, assesNo_loc2);
                    graphics.DrawString(doorNo, arialFont, Brushes.Black, doorNo_loc2);
                    graphics.DrawString(year, arialFont, Brushes.Black, year_loc2);
                    graphics.DrawString(jilla, arialFont, Brushes.Black, jilla_loc2);
                    graphics.DrawString(mandal, arialFont, Brushes.Black, mandal_loc2);
                    graphics.DrawString(name, arialFont, Brushes.Black, name_loc2);
                    graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc2);
                    graphics.DrawString(houseTax, arialFont, Brushes.Black, houseTax_loc2, format);
                    graphics.DrawString(waterTax, arialFont, Brushes.Black, waterTax_loc2, format);
                    graphics.DrawString(libSes, arialFont, Brushes.Black, libSes_loc2, format);
                    graphics.DrawString(lightTax, arialFont, Brushes.Black, lightTax_loc2, format);
                    graphics.DrawString(drainageTax, arialFont, Brushes.Black, drainageTax_loc2, format);
                    graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc2_a, format);
                    graphics.DrawString(amountInText, arialFont, Brushes.Black, amountInText_loc2);

                    //Third copy
                    PointF billNo_loc3 = new PointF(1167f, 69f);
                    PointF panchayat_loc3 = new PointF(1207f, 40f);
                    PointF assesNo_loc3 = new PointF(1167f, 110f);
                    PointF doorNo_loc3 = new PointF(1167f, 150f);
                    PointF year_loc3 = new PointF(1370f, 70f);
                    PointF jilla_loc3 = new PointF(1370f, 107f);
                    PointF mandal_loc3 = new PointF(1370f, 149f);
                    PointF name_loc3 = new PointF(1053f, 233f);
                    PointF totalTax_loc3 = new PointF(1163f, 302f);
                    PointF houseTax_loc3 = new PointF(1440f, 404f);
                    PointF waterTax_loc3 = new PointF(1440f, 423f);
                    PointF libSes_loc3 = new PointF(1440f, 448f);
                    PointF lightTax_loc3 = new PointF(1440f, 475f);
                    PointF drainageTax_loc3 = new PointF(1440f, 504f);
                    PointF totalTax_loc3_a = new PointF(1440f, 558f);
                    PointF amountInText_loc3 = new PointF(1096f, 343f);

                    graphics.DrawString(panchayat, arialFont, Brushes.Black, panchayat_loc3);
                    graphics.DrawString(billNo, arialFont, Brushes.Black, billNo_loc3);
                    graphics.DrawString(assesNum, arialFont, Brushes.Black, assesNo_loc3);
                    graphics.DrawString(doorNo, arialFont, Brushes.Black, doorNo_loc3);
                    graphics.DrawString(year, arialFont, Brushes.Black, year_loc3);
                    graphics.DrawString(jilla, arialFont, Brushes.Black, jilla_loc3);
                    graphics.DrawString(mandal, arialFont, Brushes.Black, mandal_loc3);
                    graphics.DrawString(name, arialFont, Brushes.Black, name_loc3);
                    graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc3);
                    graphics.DrawString(houseTax, arialFont, Brushes.Black, houseTax_loc3, format);
                    graphics.DrawString(waterTax, arialFont, Brushes.Black, waterTax_loc3, format);
                    graphics.DrawString(libSes, arialFont, Brushes.Black, libSes_loc3, format);
                    graphics.DrawString(lightTax, arialFont, Brushes.Black, lightTax_loc3, format);
                    graphics.DrawString(drainageTax, arialFont, Brushes.Black, drainageTax_loc3, format);
                    graphics.DrawString(totalTax, arialFont, Brushes.Black, totalTax_loc3_a, format);
                    graphics.DrawString(amountInText, arialFont, Brushes.Black, amountInText_loc3);

                }
            }

            bitmap.Save(@"C:\Users\v-kipec\Documents\demm_copy.PNG", ImageFormat.Png);
            
        }
        #endregion

    }
}
