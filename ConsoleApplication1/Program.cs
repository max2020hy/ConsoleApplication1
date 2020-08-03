using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI;
using System.IO;
using NPOI.SS.UserModel;
using static NPOI.HSSF.Util.HSSFColor;
using NPOI.HSSF.UserModel;
using Watcher;
using System.Threading;
using System.Diagnostics;

namespace ConsoleApplication1
{
    class Program
    { public static string change_File;
      public static FileSystemWatcher FSW;
        static void Main(string[] args)
        {
            //C:\Users\1\Downloads\
            string dir_old = @"c:\Users\1\Downloads\";
            if (!Directory.Exists(dir_old))
            {
                Console.WriteLine("找不到目录！");
            }
            //初始化要监视的目录
            FSW = new FileSystemWatcher(dir_old);
            //监视更改的内容
            FSW.NotifyFilter = NotifyFilters.FileName ;
            //  FSW.Filter = ".xls";
           
            FSW.Created += FSW_Created;
            FSW.EnableRaisingEvents = true;
            
            //List<Model> lst = ReadRep();

            //CreatReport(lst);            ///删除下载目录中的所有文件
            ////string[] files = Directory.GetFiles(@"c:\Users\1\Downloads");
            ////for (int i = 0; i < files.Length; i++)
            ////{

            //    File.Delete(files[i]);
            //}

            Console.ReadLine();
        }
        /// <summary>
        ///  根据模板创建新的文件
        /// </summary>
        /// <param name="lst"></param>
        private static void CreatReport(List<Model> lst)
        {
            #region 加载xls文件
            //模板文件路径
            string path = @"C:\Users\1\Desktop\日报\实收款项明细表.xls";
            FileStream fs_modle;
            using (fs_modle = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))

            {


                IWorkbook workbook_model = new NPOI.HSSF.UserModel.HSSFWorkbook(fs_modle);
               // fs_modle.Close();
               
                ISheet sheet_model = workbook_model.GetSheetAt(0);
                IRow row_hj = sheet_model.GetRow(7);
                //设置单元格时
                ICellStyle style = workbook_model.CreateCellStyle();

                style.BorderTop = BorderStyle.Thin;
                style.BorderBottom = BorderStyle.Thin;
                style.BorderLeft = BorderStyle.Thin;
                style.BorderRight = BorderStyle.Thin;
                style.WrapText = true;
                style.DataFormat = 4;
                style.VerticalAlignment = VerticalAlignment.Center;

                IRow r;

                for (int i = 6; i < lst.Count + 6; i++)
                {

                    r = sheet_model.CreateRow(i);//创建一新行

                    r.HeightInPoints = 25;//设置行高


                    for (int J = 0; J < 14; J++)
                    {
                        switch (J)
                        {
                            case 0:
                                r.CreateCell(J).SetCellValue(lst[i - 6].序号);
                                break;
                            case 1:
                                r.CreateCell(J).SetCellValue(lst[i - 6].项目名称);
                                break;
                            case 2:
                                r.CreateCell(J).SetCellValue(lst[i - 6].楼栋名称);
                                break;
                            case 3:
                                r.CreateCell(J).SetCellValue(lst[i - 6].房号);
                                break;
                            case 4:
                                r.CreateCell(J).SetCellValue(lst[i - 6].客户名称);
                                break;
                            case 5:
                                r.CreateCell(J).SetCellValue(lst[i - 6].收款日期.ToString());
                                break;
                            case 6:
                                r.CreateCell(J).SetCellValue(lst[i - 6].票据类型);
                                break;
                            case 7:
                                r.CreateCell(J).SetCellValue(lst[i - 6].票据编号);
                                break;
                            case 8:
                                r.CreateCell(J).SetCellValue(lst[i - 6].款项类型);
                                break;
                            case 9:
                                r.CreateCell(J).SetCellValue(lst[i - 6].款项名称);
                                break;
                            case 10:
                                r.CreateCell(J).SetCellValue(lst[i - 6].金额);
                                break;


                            case 11:
                                r.CreateCell(J).SetCellValue(lst[i - 6].支付方式);
                                break;
                            case 12:
                                r.CreateCell(J).SetCellValue(lst[i - 6].银付方式);
                                break;
                            case 13:
                                r.CreateCell(J).SetCellValue(lst[i - 6].摘要);
                                break;
                            default:
                                break;
                        }

                    }
                    var enu = r.GetEnumerator();

                    while (enu.MoveNext())
                    {
                        enu.Current.CellStyle = style;//设置风格
                    }


                }


                //创建合计行

                IRow r1 = sheet_model.CreateRow(lst.Count + 6);

                r1.HeightInPoints = 25;//设置行高
                ICell cell;
                IFont font = workbook_model.CreateFont();
                font.FontName = "宋体";
                font.FontHeightInPoints = 10;

                for (int i = 0; i < 14; i++)
                {

                    cell = r1.CreateCell(i);
                    cell.CellStyle = style;
                    cell.SetCellValue("--");
                    cell.CellStyle.Alignment = HorizontalAlignment.Center;

                    cell.CellStyle.SetFont(font);

                }
                double JE = lst.Sum(t => t.金额);
                r1.GetCell(0).SetCellValue(" ");
                r1.GetCell(1).SetCellValue("合计");
                r1.GetCell(10).SetCellValue(JE);

                //统计周期和统计时间
                string ZQ = $"统计周期：{DateTime.Now.ToShortDateString()} 至 {DateTime.Now.ToShortDateString()}";
                string DT = $"制表日期：{DateTime.Now.ToShortDateString()}";
                sheet_model.GetRow(3).GetCell(0).SetCellValue(ZQ);
                sheet_model.GetRow(4).GetCell(0).SetCellValue(DT);
                string newFileFullPath =$@"{ Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}";

                FileStream fs = new FileStream($"{newFileFullPath}\\日报\\日报表{DateTime.Now.ToShortDateString()}.xls", FileMode.OpenOrCreate, FileAccess.Write);
                string filePath = $"{newFileFullPath}\\日报\\日报表{DateTime.Now.ToShortDateString()}.xls";
                workbook_model.Write(fs);
                fs.Dispose();
                Console.WriteLine("文件创建成功!");
                workbook_model.Close();

                lst.Clear();
                Process.Start("wps.exe", filePath);
            }
            #endregion
        }
        /// <summary>
        /// 读取明源的报表
        /// </summary>
        /// <returns></returns>
        private static List<Model> ReadRep(string change_File)
        {
            Model model;

            int rowNum;
            List<Model> lst = new List<Model>();

            //从明源获取实收文件
         //   string file_old = @"c:\Users\1\Downloads\实收款项明细表.xls";

             
            
            change_File = change_File.Split('.')[0] + ".xls";
           



            using (FileStream fs_old = new FileStream(change_File, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook_old = new NPOI.HSSF.UserModel.HSSFWorkbook(fs_old);
                ISheet sheet_old = workbook_old.GetSheetAt(0);
            
                


            rowNum = sheet_old.LastRowNum;
            for (int i = 8; i < rowNum; i++)
            {
                model = new Model();

                model.序号 = sheet_old.GetRow(i).GetCell(1).NumericCellValue.ToString();
                model.项目名称 = sheet_old.GetRow(i).GetCell(2).StringCellValue;
                model.楼栋名称 = sheet_old.GetRow(i).GetCell(5).StringCellValue;
                // CellType ct = sheet_old.GetRow(i).GetCell(14).CellType;
                model.房号 = sheet_old.GetRow(i).GetCell(9).StringCellValue;
                model.客户名称 = sheet_old.GetRow(i).GetCell(11).StringCellValue;
                model.收款日期 = sheet_old.GetRow(i).GetCell(14).DateCellValue;
                model.票据类型 = sheet_old.GetRow(i).GetCell(17).StringCellValue;
                model.票据编号 = sheet_old.GetRow(i).GetCell(18).StringCellValue;
                model.款项类型 = sheet_old.GetRow(i).GetCell(19).StringCellValue;
                model.款项名称 = sheet_old.GetRow(i).GetCell(20).StringCellValue;
                model.金额 = sheet_old.GetRow(i).GetCell(22).NumericCellValue;
                model.支付方式 = sheet_old.GetRow(i).GetCell(25).StringCellValue;
                model.银付方式 = sheet_old.GetRow(i).GetCell(27).StringCellValue;
                model.摘要 = sheet_old.GetRow(i).GetCell(26).StringCellValue;
                lst.Add(model);


            }
          
            workbook_old.Close();
             }
            return lst;
        }
          public static List<Model> Lst;
        private static void FSW_Created(object sender, FileSystemEventArgs e)
        {
            
            Thread.Sleep(3000);
           // change_File = e.FullPath;
            Console.WriteLine($"文件{e.FullPath}被创建");
          //  FSW.EnableRaisingEvents = false;
            Lst= ReadRep(e.FullPath);
            CreatReport(Lst);
            Lst.Clear();
        }
    }
}
