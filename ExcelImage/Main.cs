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

namespace LYF.ExcelImage
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void btNewFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Excel2003文件";
            sfd.FileName = "";
            sfd.Filter = "Excel2003文件(*.xls)|*.xls";
            if(sfd.ShowDialog()==DialogResult.OK)
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                ISheet sheet1 = workbook.CreateSheet("FQC");

                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "pkm";
                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Title =
                si.Subject = "automatic genereted document";
                si.Author = "pkm";
                workbook.DocumentSummaryInformation = dsi;
                workbook.SummaryInformation = si;

                ICellStyle titleCellStyle = workbook.CreateCellStyle();
                titleCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                titleCellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                IFont font = workbook.CreateFont();
                font.FontHeightInPoints = 16;
                font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
                font.FontName = "微软雅黑";
                titleCellStyle.SetFont(font);//HEAD 样式

                IRow row0 = sheet1.CreateRow(0);
                NPOI.SS.UserModel.ICell titleCell = row0.CreateCell(0);
                titleCell.SetCellValue("订单测试表头啊");
                titleCell.CellStyle = titleCellStyle;
                sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));

                //CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 10);
                //sheet1.AddMergedRegion(cellRangeAddress);



                //excel格式化
                ICellStyle dateStyle = workbook.CreateCellStyle();
                dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy/m/d h:mm:ss");

                ICellStyle numberStyle = workbook.CreateCellStyle();
                numberStyle.DataFormat = workbook.CreateDataFormat().GetFormat("0.00000");

                ICellStyle textStyle = workbook.CreateCellStyle();
                textStyle.DataFormat = workbook.CreateDataFormat().GetFormat("@");

                //给sheet1添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(1);
                row1.CreateCell(0).SetCellValue("订单号");
                row1.CreateCell(1).SetCellValue("条码");
                row1.CreateCell(2).SetCellValue("档位名称");
                row1.CreateCell(3).SetCellValue("Pmax");
                row1.CreateCell(4).SetCellValue("功率档");
                row1.CreateCell(5).SetCellValue("功率档范围");
                row1.CreateCell(6).SetCellValue("Ipm");
                row1.CreateCell(7).SetCellValue("电流档");
                row1.CreateCell(8).SetCellValue("电流档范围");
                row1.CreateCell(9).SetCellValue("规格");
                row1.CreateCell(10).SetCellValue("产品等级");

                for (int i = 0; i < 100; i++)
                {
                    IRow row2 = sheet1.CreateRow(i+2);
                    NPOI.SS.UserModel.ICell cell0 = row2.CreateCell(0);
                    cell0.SetCellValue("订单号"+i);

                    ICellStyle style = workbook.CreateCellStyle();
                    //设置单元格的样式：水平对齐居中
                    style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    //新建一个字体样式对象
                    IFont cellFont = workbook.CreateFont();
                    //设置字体加粗样式
                    font.Boldweight = short.MaxValue;
                    //使用SetFont方法将字体样式添加到单元格样式中 
                    style.SetFont(cellFont);


                    style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

                    //将新的样式赋给单元格
                    cell0.CellStyle = style;

                    row2.CreateCell(1).SetCellValue("条码" + i);
                    row2.CreateCell(2).SetCellValue("档位名称" + i);
                    row2.CreateCell(3).SetCellValue("Pmax" + i);
                    row2.CreateCell(4).SetCellValue("功率档" + i);
                    row2.CreateCell(5).SetCellValue("功率档范围" + i);
                    row2.CreateCell(6).SetCellValue("Ipm" + i);
                    row2.CreateCell(7).SetCellValue("电流档" + i);
                    row2.CreateCell(8).SetCellValue("电流档范围" + i);
                    row2.CreateCell(9).SetCellValue("规格" + i);
                    row2.CreateCell(10).SetCellValue("产品等级" + i);
                }

                System.IO.MemoryStream ms = new System.IO.MemoryStream();
                workbook.Write(ms);
                File.WriteAllBytes(sfd.FileName, ms.ToArray());
                MessageBox.Show("OK");

            }
        }

        private void btNewFile2007_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Excel2007文件";
            sfd.FileName = "";
            sfd.Filter = "Excel2007文件(*.xlsx)|*.xlsx";
            if (sfd.ShowDialog() == DialogResult.OK)
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
                ICellStyle dateStyle = workbook.CreateCellStyle();
                dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy/m/d h:mm:ss");

                ICellStyle numberStyle = workbook.CreateCellStyle();
                numberStyle.DataFormat = workbook.CreateDataFormat().GetFormat("0.00000");

                ICellStyle textStyle = workbook.CreateCellStyle();
                textStyle.DataFormat = workbook.CreateDataFormat().GetFormat("@");

                //给sheet1添加第一行的头部标题
                IRow row1 = sheet1.CreateRow(0);
                row1.CreateCell(0).SetCellValue("订单号");
                row1.CreateCell(1).SetCellValue("条码");
                row1.CreateCell(2).SetCellValue("档位名称");
                row1.CreateCell(3).SetCellValue("Pmax");
                row1.CreateCell(4).SetCellValue("功率档");
                row1.CreateCell(5).SetCellValue("功率档范围");
                row1.CreateCell(6).SetCellValue("Ipm");
                row1.CreateCell(7).SetCellValue("电流档");
                row1.CreateCell(8).SetCellValue("电流档范围");
                row1.CreateCell(9).SetCellValue("规格");
                row1.CreateCell(10).SetCellValue("产品等级");

                for (int i = 0; i < 100; i++)
                {
                    IRow row0 = sheet1.CreateRow(i+1);
                    row0.CreateCell(0).SetCellValue("订单号" + i);
                    row0.CreateCell(1).SetCellValue("条码" + i);
                    row0.CreateCell(2).SetCellValue("档位名称" + i);
                    row0.CreateCell(3).SetCellValue("Pmax" + i);
                    row0.CreateCell(4).SetCellValue("功率档" + i);
                    row0.CreateCell(5).SetCellValue("功率档范围" + i);
                    row0.CreateCell(6).SetCellValue("Ipm" + i);
                    row0.CreateCell(7).SetCellValue("电流档" + i);
                    row0.CreateCell(8).SetCellValue("电流档范围" + i);
                    row0.CreateCell(9).SetCellValue("规格" + i);
                    row0.CreateCell(10).SetCellValue("产品等级" + i);
                }

                byte[] bytes = System.IO.File.ReadAllBytes(Path.Combine(Application.StartupPath,"img.jpg"));
                int pictureIdx = workbook.AddPicture(bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);

                #region 写入图片
                // Create the drawing patriarch.  This is the top level container for all shapes. 
                IDrawing patriarch = sheet1.CreateDrawingPatriarch();
                //add a picture
                IClientAnchor anchor = new XSSFClientAnchor(0, 0, 1023, 0, 0, 0, 1, 3);
                IPicture pict = patriarch.CreatePicture(anchor, pictureIdx);
                pict.Resize();
                #endregion

                System.IO.MemoryStream ms = new System.IO.MemoryStream();
                workbook.Write(ms);
                File.WriteAllBytes(sfd.FileName, ms.ToArray());
                MessageBox.Show("OK");

            }
        }


        #region 代码备用

        #region 添加图片
        //byte[] bytes = System.IO.File.ReadAllBytes(@"D:\MyProject\NPOIDemo\ShapeImage\image1.jpg");
        //int pictureIdx = hssfworkbook.AddPicture(bytes, HSSFWorkbook.PICTURE_TYPE_JPEG);

        ////create sheet
        //HSSFSheet sheet = hssfworkbook.CreateSheet("Sheet1");

        //// Create the drawing patriarch.  This is the top level container for all shapes. 
        //HSSFPatriarch patriarch = sheet.CreateDrawingPatriarch();

        ////add a picture
        //HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 0, 0, 0, 1, 3);
        //HSSFPicture pict = patriarch.CreatePicture(anchor, pictureIdx);
        //pict.Resize();
        #endregion

        #endregion

        private void btNewWord2003_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Word2003文件";
            sfd.FileName = "";
            sfd.Filter = "Word2003文件(*.doc)|*.doc";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                
                //NPOI.HWPF.HWPFDocument
                using (MemoryStream ms = new MemoryStream())
                {
                    HWPFDocument hd = new HWPFDocument(ms);
                    Range rang = hd.GetRange();
                    rang.InsertAfter("测试一下把！");
                    //var table = hd.ParagraphTable;
                    //ListTables ta=  hd.GetListTables();
                    //TextPieceTable tp = hd.TextTable;
                    //Range rang = hd.GetRange();
                    //int paragraphCount = rang.NumParagraphs;
                    //for (int i = 0; i < paragraphCount; i++)
                    //{
                    //    CharacterRun pph = rang.GetCharacterRun(i);
                    //    sb.Append(pph.Text);
                    //}
                    hd.Write(ms);
                    SaveToFile(ms, sfd.FileName);
                }


                //m_Docx = CreatDocxTable();
                //m_Docx.Write(ms);
                //ms.Flush();
                //SaveToFile(ms, sfd.FileName);
                MessageBox.Show("保存成功!");
            }
        }

        private void btNewWord2007_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Word2007文件";
            sfd.FileName = "";
            sfd.Filter = "Word2007文件(*.docx)|*.docx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                MemoryStream ms = new MemoryStream();
                XWPFDocument m_Docx = new XWPFDocument();
                m_Docx = CreatDocxTable();
                m_Docx.Write(ms);
                ms.Flush();
                SaveToFile(ms, sfd.FileName);
                MessageBox.Show("保存成功!");
            }
        }

        protected XWPFDocument CreatDocxTable()
        {
            XWPFDocument m_Docx = new XWPFDocument();
            XWPFParagraph p0 = m_Docx.CreateParagraph();
            XWPFRun r0 = p0.CreateRun();
            r0.SetText("DOCX表");

            XWPFTable table = m_Docx.CreateTable(1, 3);//创建一行3列表
            table.GetRow(0).GetCell(0).SetText("111");
            table.GetRow(0).GetCell(1).SetText("222");
            table.GetRow(0).GetCell(2).SetText("333");

            XWPFTableRow m_Row = table.CreateRow();//创建一行
            m_Row = table.CreateRow();//创建一行
            m_Row.GetCell(0).SetText("211");

            //合并单元格
            m_Row = table.InsertNewTableRow(0);//表头插入一行
            XWPFTableCell cell = m_Row.CreateCell();//创建一个单元格,创建单元格时就创建了一个CT_P
            CT_Tc cttc = cell.GetCTTc();
            CT_TcPr ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan = new CT_DecimalNumber();
            ctPr.gridSpan.val = "3";//合并3列
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "abc";

            XWPFTableRow td3 = table.InsertNewTableRow(table.Rows.Count - 1);//插入行
            cell = td3.CreateCell();
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan = new CT_DecimalNumber();
            ctPr.gridSpan.val = "3";
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "qqq";

            //表增加行，合并列
            CT_Row m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row); //必须要！！！
            cell = m_Row.CreateCell();
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan = new CT_DecimalNumber();
            ctPr.gridSpan.val = "3";
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "sss";

            //表未增加行，合并2列，合并2行
            //1行
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan = new CT_DecimalNumber();
            ctPr.gridSpan.val = "2";
            ctPr.AddNewVMerge().val = ST_Merge.restart;//合并行
            ctPr.AddNewVAlign().val = ST_VerticalJc.center;//垂直居中
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "xxx";
            cell = m_Row.CreateCell();
            cell.SetText("ddd");
            //2行，多行合并类似
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan = new CT_DecimalNumber();
            ctPr.gridSpan.val = "2";
            ctPr.AddNewVMerge().val = ST_Merge.@continue;//合并行
            cell = m_Row.CreateCell();
            cell.SetText("kkk");
            ////3行
            //m_NewRow = new CT_Row();
            //m_Row = new XWPFTableRow(m_NewRow, table);
            //table.AddRow(m_Row);
            //cell = m_Row.CreateCell();
            //cttc = cell.GetCTTc();
            //ctPr = cttc.AddNewTcPr();
            //ctPr.gridSpan.val = "2";
            //ctPr.AddNewVMerge().val = ST_Merge.@continue;
            //cell = m_Row.CreateCell();
            //cell.SetText("hhh");

            return m_Docx;
        }
        static void SaveToFile(MemoryStream ms, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();

                fs.Write(data, 0, data.Length);
                fs.Flush();
                data = null;
            }
        }

        private void btExcelImage_Click(object sender, EventArgs e)
        {
            GetExcelImage();
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
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "jpg|*.jpg|bmp|*.bmp|png|*.png";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Bitmap bmp = new Bitmap(Image.FromFile(ofd.FileName));


                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Title = "Excel2007文件";
                sfd.FileName = "";
                sfd.Filter = "Excel2007文件(*.xlsx)|*.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
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
                    ICellStyle dateStyle = workbook.CreateCellStyle();
                    dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy/m/d h:mm:ss");

                    ICellStyle numberStyle = workbook.CreateCellStyle();
                    numberStyle.DataFormat = workbook.CreateDataFormat().GetFormat("0.00000");

                    ICellStyle textStyle = workbook.CreateCellStyle();
                    textStyle.DataFormat = workbook.CreateDataFormat().GetFormat("@");

                    //CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 10);
                    //sheet1.AddMergedRegion(cellRangeAddress);

                    for (int i = 0; i < bmp.Height; i++)
                    {
                        IRow row0 = sheet1.CreateRow(i);
                        row0.HeightInPoints = 3;

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
                        for (int j = 0; j < bmp.Width; j++)
                        {
                            NPOI.SS.UserModel.ICell cell = row0.CreateCell(j);
                            ICellStyle style = workbook.CreateCellStyle();
                            ////设置单元格的样式：水平对齐居中
                            //style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                            ////新建一个字体样式对象
                            //IFont cellFont = workbook.CreateFont();
                            ////设置字体加粗样式
                            ////font.Boldweight = short.MaxValue;
                            ////使用SetFont方法将字体样式添加到单元格样式中 
                            //style.SetFont(cellFont);


                            //style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                            //style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                            //style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                            //style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
     
                            //HSSFColor color=new HSSFColor();
                            //color.
                            //style.FillBackgroundColor= GetXLColour(workbook,bmp.GetPixel(i,j));
                            Color color = bmp.GetPixel(j,i);

                            //XSSFColor xssfColor = new XSSFColor();
                            ////根据自己需要设置RGB
                            //byte[] colorRgb = { (byte)color.R, (byte)color.G, (byte)color.B };
                            //xssfColor.SetRgb(colorRgb);
                            //style.FillBackgroundColorColor = xssfColor;
                            //style.FillPattern = FillPattern.SolidForeground;
                            //style.FillForegroundColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(color.R, color.G, color.B).Indexed; 03

                            style.FillForegroundColor= 1;
                            ((XSSFColor)style.FillForegroundColorColor).SetRgb(new byte[] { color.R, color.G, color.B });
                            cell.SetCellValue("0");
                            cell.CellStyle = style;
                        }
                    }
                    //设置列宽
                    for (int j = 0; j < bmp.Width; j++)
                    {
                        sheet1.SetColumnWidth(j, 1);  
                    }

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
                    File.WriteAllBytes(sfd.FileName, ms.ToArray());
                    MessageBox.Show("OK");

                }
            }
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
    }
}
