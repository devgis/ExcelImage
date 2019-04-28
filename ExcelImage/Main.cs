using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.HSSF.Util;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace LYF.ExcelImage
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        Bitmap bmp = null;
        int bmpwidth = 0;
        int bmpheight = 0;
        string savefilename = string.Empty;
        Thread trun=null;
        private void btExcelImage_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "jpg|*.jpg|bmp|*.bmp|png|*.png";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                bmp = new Bitmap(Image.FromFile(ofd.FileName));
                //bmpwidth = bmp.Width;
                //bmpheight = bmp.Height;

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Title = "Excel2007文件";
                sfd.FileName = "";
                sfd.Filter = "Excel2007文件(*.xlsx)|*.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    savefilename = sfd.FileName;
                    progressBar.Value = 0;
                    trun = new Thread(GetExcelImage);
                    trun.Start();
                }
            }
            
            //try
            //{
            //    GetExcelImage();
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show("失败："+ex.Message);
            //}
        }

        private void GetExcelImage()
        { 

                    XSSFWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet1 = workbook.CreateSheet("FQC");
                    var coreProps = workbook.GetProperties(); //获取文件属性信息
                    coreProps.CoreProperties.Category = "默认分类";

                    coreProps.CoreProperties.Creator = "Liyafei";
                    coreProps.CoreProperties.Description = "李亚飞写的文档！";
                    coreProps.CoreProperties.Subject = "Subject";
                    coreProps.CustomProperties.AddProperty("LaYafei自定义属性", "自定义属性内容");
                    //coreProps.ExtendedProperties = "LiyafeiApp1.0"; //只能读取的属性

                    //excel格式化
                    //ICellStyle dateStyle = workbook.CreateCellStyle();
                    //dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy/m/d h:mm:ss");

                    //ICellStyle numberStyle = workbook.CreateCellStyle();
                    //numberStyle.DataFormat = workbook.CreateDataFormat().GetFormat("0.00000");

                    //ICellStyle textStyle = workbook.CreateCellStyle();
                    //textStyle.DataFormat = workbook.CreateDataFormat().GetFormat("@");

                    //CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 10);
                    //sheet1.AddMergedRegion(cellRangeAddress);
                    //for (int i = 0; i < bmp.Height; i++)
                    //{
                    //    IRow row0 = sheet1.CreateRow(i);
                    //    for (int j = 0; j < bmp.Width; j++)
                    //    {
                    //        row0.CreateCell(j);
                    //    }
                    //}
                    lbInfo.Invoke((EventHandler)(delegate
                    {
                        lbInfo.Text = "开始创建表格！";
                    }));
                    for (int i = 0; i < bmp.Height; i++)
                    {
                        IRow row0 = sheet1.CreateRow(i);
                        row0.HeightInPoints = 50;

                        //row2.CreateCell(1).SetCellValue("条码" + i);
                        //row2.CreateCell(2).SetCellValue("档位名称" + i);
                        //row2.CreateCell(3).SetCellValue("Pmax" + i);
                        //row2.CreateCell(4).SetCellValue("功率档" + i);
                        //row2.CreateCell(5).SetCellValue("功率档范围" + i);
                        //row2.CreateCell(6).SetCellValue("Ipm" + i);
                        //row2.CreateCell(7).SetCellValue("电流档" + i);
                        //row2.CreateCell(8).SetCellValue("电流档范围" + i);
                        //row2.CreateCell(9).SetCellValue("规格" + i);
                        //row2.CreateCell(10).SetCellValue("产品等级" + i);

                        //ThreadPool.QueueUserWorkItem(o =>
                        //{
                        //    int row = (int)o;

                        //},i);
                        
                        for (int j = 0; j < bmp.Width; j++)
                        {
                            row0.CreateCell(j);
                        }
                    }
                    lbInfo.Invoke((EventHandler)(delegate
                    {
                        lbInfo.Text = "表格创建完毕！";
                    }));
       
                    for (int i = 0; i < bmp.Height; i++)
                    {
                        IRow row0 = sheet1.GetRow(i);
                        row0.HeightInPoints = 50;

                        //row2.CreateCell(1).SetCellValue("条码" + i);
                        //row2.CreateCell(2).SetCellValue("档位名称" + i);
                        //row2.CreateCell(3).SetCellValue("Pmax" + i);
                        //row2.CreateCell(4).SetCellValue("功率档" + i);
                        //row2.CreateCell(5).SetCellValue("功率档范围" + i);
                        //row2.CreateCell(6).SetCellValue("Ipm" + i);
                        //row2.CreateCell(7).SetCellValue("电流档" + i);
                        //row2.CreateCell(8).SetCellValue("电流档范围" + i);
                        //row2.CreateCell(9).SetCellValue("规格" + i);
                        //row2.CreateCell(10).SetCellValue("产品等级" + i);
                    
                            //ThreadPool.QueueUserWorkItem(o =>
                            //{
                            //    int row = (int)o;
                               
                            //},i);
                        //row0.CreateCell(bmp.Width);
                        for (int j = 0; j < bmp.Width; j++)
                        {
                            ICellStyle style = workbook.CreateCellStyle();
                            style.FillPattern = FillPattern.SolidForeground;
                            //style.FillBackgroundColor = 12;
                            //style.FillForegroundColor = 12;
                            //style.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
                            Color color = bmp.GetPixel(j, i);
                            style.FillForegroundColor = 1;
                            ((XSSFColor)style.FillForegroundColorColor).SetRgb(new byte[] { color.R, color.G, color.B });
                            row0.GetCell(j).CellStyle = style;

                            progressBar.Invoke((EventHandler)(delegate
                            {
                                progressBar.Value = (i * bmp.Width + j) *100/ (bmp.Width * bmp.Height);
                                lbInfo.Invoke((EventHandler)(delegate
                                {
                                    lbInfo.Text = string.Format("{0}%", progressBar.Value);
                                }));
                                //progressBar.Refresh();

                            }));
                            //NPOI.SS.UserModel.ICell cell = row0.CreateCell(j,CellType.String);
                        }
                            //row0.CreateCell(j);//.SetCellValue("");

                            //ICellStyle style = workbook.CreateCellStyle();
                            //style.FillPattern = FillPattern.SolidForeground;
                            ////style.FillBackgroundColor = 12;
                            ////style.FillForegroundColor = 12;
                            ////style.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
                            //Color color = bmp.GetPixel(j,i);
                            //style.FillForegroundColor = 1;
                            //((XSSFColor)style.FillForegroundColorColor).SetRgb(new byte[] { color.R, color.G, color.B });
                            //row0.CreateCell(j).CellStyle = style;
                            ////NPOI.SS.UserModel.ICell cell = row0.CreateCell(j,CellType.String);


                            //ICellStyle style = workbook.CreateCellStyle();
                            //////设置单元格的样式：水平对齐居中
                            ////style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                            //////新建一个字体样式对象
                            ////IFont cellFont = workbook.CreateFont();
                            //////设置字体加粗样式
                            //////font.Boldweight = short.MaxValue;
                            //////使用SetFont方法将字体样式添加到单元格样式中 
                            ////style.SetFont(cellFont);


                            ////style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                            ////style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                            ////style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                            ////style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
     
                            ////HSSFColor color=new HSSFColor();
                            ////color.
                            ////style.FillBackgroundColor= GetXLColour(workbook,bmp.GetPixel(i,j));
                            //Color color = bmp.GetPixel(j,i);

                            ////XSSFColor xssfColor = new XSSFColor();
                            //////根据自己需要设置RGB
                            ////byte[] colorRgb = { (byte)color.R, (byte)color.G, (byte)color.B };
                            ////xssfColor.SetRgb(colorRgb);
                            ////style.FillBackgroundColorColor = xssfColor;
                            ////style.FillPattern = FillPattern.SolidForeground;
                            ////style.FillForegroundColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(color.R, color.G, color.B).Indexed; 03

                            ////style.FillForegroundColor= 1;
                            ////((XSSFColor)style.FillForegroundColorColor).SetRgb(new byte[] { color.R, color.G, color.B });
                            ////cell.SetCellValue("0");
                            ////cell.CellStyle = style;

                            //style.FillPattern = FillPattern.SolidForeground;
                            //style.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index;
                            //style.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;

                            ////style.FillForegroundColor = HSSFColor.Pink.Index;
                            ////style.FillPattern = FillPattern.Squares;

                            //row0.Cells[j].CellStyle = style;
                            //row0.Cells[j].SetCellValue("吃法");
                        }
                    //}
                    //设置列宽
                    //for (int j = 0; j < bmp.Width; j++)
                    //{
                    //    sheet1.SetColumnWidth(j, 1);  
                    //}

                    //for (int i = 0; i < bmp.Height; i++)
                    //{
                    //    IRow row2 = sheet1.CreateRow(i);


                    //    //row2.CreateCell(1).SetCellValue("条码" + i);
                    //    //row2.CreateCell(2).SetCellValue("档位名称" + i);
                    //    //row2.CreateCell(3).SetCellValue("Pmax" + i);
                    //    //row2.CreateCell(4).SetCellValue("功率档" + i);
                    //    //row2.CreateCell(5).SetCellValue("功率档范围" + i);
                    //    //row2.CreateCell(6).SetCellValue("Ipm" + i);
                    //    //row2.CreateCell(7).SetCellValue("电流档" + i);
                    //    //row2.CreateCell(8).SetCellValue("电流档范围" + i);
                    //    //row2.CreateCell(9).SetCellValue("规格" + i);
                    //    //row2.CreateCell(10).SetCellValue("产品等级" + i);
                    //    for (int j = 0; j < bmp.Width; j++)
                    //    {
                    //        NPOI.SS.UserModel.ICell cell = row2.CreateCell(j);
                    //        ICellStyle style = workbook.CreateCellStyle();
                    //        //设置单元格的样式：水平对齐居中
                    //        style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    //        //新建一个字体样式对象
                    //        IFont cellFont = workbook.CreateFont();
                    //        //设置字体加粗样式
                    //        font.Boldweight = short.MaxValue;
                    //        //使用SetFont方法将字体样式添加到单元格样式中 
                    //        style.SetFont(cellFont);


                    //        style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    //        style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    //        style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    //        style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    //        //HSSFColor color=new HSSFColor();
                    //        //color.
                    //        //style.FillBackgroundColor= GetXLColour(workbook,bmp.GetPixel(i,j));
                    //        Color color = bmp.GetPixel(i, j);

                    //        XSSFColor xssfColor = new XSSFColor();
                    //        //根据自己需要设置RGB
                    //        byte[] colorRgb = { (byte)color.R, (byte)color.G, (byte)color.B };
                    //        xssfColor.SetRgb(colorRgb);
                    //        //style.FillBackgroundColorColor = xssfColor;
                    //        //style.FillPattern = FillPattern.SolidForeground;
                    //        style.FillForegroundColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(color.R, color.G, color.B).Indexed;

                    //        cell.CellStyle = style;
                    //    }
                    //}

                    System.IO.MemoryStream ms = new System.IO.MemoryStream();
                    workbook.Write(ms);
                    File.WriteAllBytes(savefilename, ms.ToArray());
                    progressBar.Invoke((EventHandler)(delegate
                    {
                        progressBar.Value = 100;
                        //progressBar.Refresh();

                    }));
                    lbInfo.Invoke((EventHandler)(delegate
                    {
                        lbInfo.Text = "完成";
                    }));
        }

        private short GetXLColour(HSSFWorkbook workbook, System.Drawing.Color SystemColour)
        {
            short s = 0;
            HSSFPalette XlPalette = workbook.GetCustomPalette();
            HSSFColor XlColour = XlPalette.FindColor(SystemColour.R, SystemColour.G, SystemColour.B);
            if (XlColour == null)
            {
                XlColour = XlPalette.AddColor(SystemColour.R, SystemColour.G, SystemColour.B);
                s = XlColour.Indexed;
                //if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 255)
                //{
                //    if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 64)
                //    {
                //        //NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE = (byte)64;
                //        //NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE += 1;
                //        XlColour = XlPalette.AddColor(SystemColour.R, SystemColour.G, SystemColour.B);
                //    }
                //    else
                //    {
                //        XlColour = XlPalette.FindSimilarColor(SystemColour.R, SystemColour.G, SystemColour.B);
                //    }

                //    s = XlColour.Indexed;
                //}

            }
            else
                s = XlColour.Indexed;

            return s;
        }

        private void Main_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (trun != null)
            {
                try
                {
                    trun.Abort();
                    trun.Join();
                }
                catch
                { }
            }
        }

        private void btExcelOffice_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "jpg|*.jpg|bmp|*.bmp|png|*.png";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                bmp = new Bitmap(Image.FromFile(ofd.FileName));
                //bmpwidth = bmp.Width;
                //bmpheight = bmp.Height;

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Title = "Excel2007文件";
                sfd.FileName = "";
                sfd.Filter = "Excel2007文件(*.xlsx)|*.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    btExcelOffice.Enabled = false;
                    savefilename = sfd.FileName;
                    progressBar.Value = 0;
                    trun = new Thread(GetExcelImageOffice);
                    trun.Start();
                }
            }
            
        }


        private void GetExcelImageOffice()
        {
            try
            {
                Microsoft.Office.Interop.Excel.ApplicationClass myExcel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                myExcel.Visible = false;
                myExcel.DisplayAlerts = false;  //是否需要显示提示
                myExcel.AlertBeforeOverwriting = false;  //是否弹出提示覆盖

                myExcel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                //停用警告訊息
                myExcel.DisplayAlerts = false;
                myExcel.UserControl = true;
                //讓Excel文件可見
                myExcel.Visible = false;
                //引用第一個活頁簿
                Workbook myBook = myExcel.Workbooks[1];
                //設定活頁簿焦點
                Worksheet sheet = myBook.ActiveSheet as Worksheet;
                sheet.Name = "图片";

                lbInfo.Invoke((EventHandler)(delegate
                {
                    lbInfo.Text = "开始创建表格！";
                }));
                for (int i = 0; i < bmp.Height; i++)
                {
                    for (int j = 0; j < bmp.Width; j++)
                    {
                        Color color = bmp.GetPixel(j, i);

                        Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[i + 1, j + 1];
                        range.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(color.R, color.G, color.B));//设置颜色
                        range.ColumnWidth = 0.3;
                        range.RowHeight = 3;
                        progressBar.Invoke((EventHandler)(delegate
                        {
                            progressBar.Value = (i * bmp.Width + j) * 100 / (bmp.Width * bmp.Height);
                            lbInfo.Invoke((EventHandler)(delegate
                            {
                                lbInfo.Text = string.Format("{0}%", progressBar.Value);
                            }));
                        }));
                    }
                }
                myBook.SaveAs(savefilename);

                myBook.Close(true, Type.Missing, Type.Missing);
                myBook = null;
                myExcel.Quit();
                myExcel = null;
                btExcelOffice.Invoke((EventHandler)(delegate
                {
                    btExcelOffice.Enabled = true;
                }));
                lbInfo.Invoke((EventHandler)(delegate
                {
                    lbInfo.Text = "完成";
                }));
            }
            catch(Exception ex)
            {
                lbInfo.Invoke((EventHandler)(delegate
                {
                    lbInfo.Text = "发生错误:"+ex.Message;
                }));
                btExcelOffice.Invoke((EventHandler)(delegate
                {
                    btExcelOffice.Enabled = true;
                }));
            }

            
        }

        private void btStop_Click(object sender, EventArgs e)
        {
            try
            {
                trun.Abort();
                trun.Join();
            }
            catch
            { }
            lbInfo.Text = "就绪!";
            btExcelOffice.Enabled = true;
        }
    }
}
