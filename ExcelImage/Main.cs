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
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.HSSF.Util;
using System.Threading;

namespace LYF.ExcelImage
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        Bitmap bmp = null;
        string savefilename = string.Empty;
        Thread trun=null;
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
                //2013只支持256列  所以需要用新版
                XSSFWorkbook wb = new XSSFWorkbook();
                XSSFSheet sheet = (XSSFSheet)wb.CreateSheet("Image");

                lbInfo.Invoke((EventHandler)(delegate
                {
                    lbInfo.Text = "开始创建表格！";
                }));

               
                for (int i = 0; i < bmp.Height; i++)
                {
                    XSSFRow imageRow = (XSSFRow)sheet.CreateRow(i);
                    imageRow.Height = 5;
                    for (int j = 0; j < bmp.Width; j++)
                    {
                        XSSFCell headCell = (XSSFCell)imageRow.CreateCell(j,CellType.String);
                        XSSFColor xssfColor = new XSSFColor(bmp.GetPixel(j, i));
                        XSSFCellStyle colorStyle = (XSSFCellStyle)wb.CreateCellStyle();
                        colorStyle.FillBackgroundXSSFColor = xssfColor;
                        headCell.CellStyle = colorStyle;
                      
                        ThreadPool.QueueUserWorkItem(o =>
                        {
                            
                            if (o != null && o is int)
                            {
                                progressBar.Invoke((EventHandler)(delegate
                                {
                                    progressBar.Value = (int)o;
                                    lbInfo.Invoke((EventHandler)(delegate
                                    {
                                        lbInfo.Text = string.Format("{0}%", progressBar.Value);
                                    }));
                                }));
                            }

                        }, (i * bmp.Width + j) * 100 / (bmp.Width * bmp.Height));
                    }
                }

                for (int i = 0; i < bmp.Height; i++)
                {
                    sheet.SetColumnWidth(i, 5);
                }
                using (FileStream fs = new FileStream(savefilename, FileMode.Create))
                {
                    wb.Write(fs);
                }
                btExcelOffice.Invoke((EventHandler)(delegate
                {
                    btExcelOffice.Enabled = true;
                }));
                lbInfo.Invoke((EventHandler)(delegate
                {
                    lbInfo.Text = "完成";
                }));
                progressBar.Invoke((EventHandler)(delegate
                {
                    progressBar.Value = 0;
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
                progressBar.Invoke((EventHandler)(delegate
                {
                    progressBar.Value = 0;
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
            progressBar.Invoke((EventHandler)(delegate
            {
                progressBar.Value = 0;
            }));
        }

    }

    struct MyColor
    {
        public byte R;
        public byte G;
        public byte B;
    }
}
