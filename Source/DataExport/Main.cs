using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace DataExport
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void btImport_Click(object sender, EventArgs e)
        {
            //开始导入
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Excel文件";
            ofd.FileName = "";
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);//为了获取特定的系统文件夹，可以使用System.Environment类的静态方法GetFolderPath()。该方法接受一个Environment.SpecialFolder枚举，其中可以定义要返回路径的哪个系统目录
            ofd.Filter = "Excel文件(*.xlsx)|*.xlsx";
            ofd.ValidateNames = true;     //文件有效性验证ValidateNames，验证用户输入是否是一个有效的Windows文件名
            ofd.CheckFileExists = true;  //验证路径有效性
            ofd.CheckPathExists = true; //验证文件有效性
            string strName = string.Empty;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                tbImportPath.Text = ofd.FileName;
            }
        }

        private void btExrpot_Click(object sender, EventArgs e)
        {
            if (!File.Exists(tbImportPath.Text))
            {
                MessageBox.Show("导入文件不存在！");
                return;
            }
            

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "Excel文件";
            sfd.FileName = "";
            sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);//为了获取特定的系统文件夹，可以使用System.Environment类的静态方法GetFolderPath()。该方法接受一个Environment.SpecialFolder枚举，其中可以定义要返回路径的哪个系统目录
            sfd.Filter = "Excel文件(*.xlsx)|*.xlsx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                this.tbExportPath.Text = sfd.FileName;

                //读数据
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();//lauch excel application
                if (excel == null)
                {
                    MessageBox.Show("未安装Excel");
                }
                else
                {
                    excel.Visible = false;
                    excel.UserControl = true;
                    // 以只读的形式打开EXCEL文件
                    Workbook wb = excel.Application.Workbooks.Open(this.tbImportPath.Text, missing, true, missing, missing, missing,
                     missing, missing, missing, true, missing, missing, missing, missing, missing);
                    //取得第一个工作薄
                    Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);

                    Workbook exportWorkbook = excel.Workbooks.Add(true);
                    Worksheet exportWorksheet = (Worksheet)exportWorkbook.ActiveSheet;

                    //取得总记录行数   (包括标题列)
                    int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                    //Range rng1 = ws.Cells.get_Range("A5", "A" + rowsint);   //item
                    if (rowsint <= 5)
                    {
                        MessageBox.Show("没有数据或者数据格式不正确！");
                        return;
                    }
                    else
                    {

                        //object[,] oA = (object[,])ws.get_Range(ws.Cells[5, 1], ws.Cells[rowsint, 1]).Value2; //来料型号 

                        exportWorksheet.Activate();
                        exportWorksheet.Cells[1, 1] = "XXXXXXX股份有限公司";
                        exportWorksheet.Cells[2, 1] = "CP制令单";

                        //exportWorksheet.Cells.get_Range("A1", "U1").EntireColumn.AutoFit();//自动调整列宽
                        //设置列宽
                        ((Range)exportWorksheet.Cells[1, 1]).EntireColumn.ColumnWidth = 10;
                        ((Range)exportWorksheet.Cells[1, 2]).EntireColumn.ColumnWidth = 10;
                        ((Range)exportWorksheet.Cells[1, 3]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 4]).EntireColumn.ColumnWidth = 10;
                        ((Range)exportWorksheet.Cells[1, 5]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 6]).EntireColumn.ColumnWidth = 6;
                        ((Range)exportWorksheet.Cells[1, 7]).EntireColumn.ColumnWidth = 10;
                        ((Range)exportWorksheet.Cells[1, 8]).EntireColumn.ColumnWidth = 10;
                        ((Range)exportWorksheet.Cells[1, 9]).EntireColumn.ColumnWidth = 10;
                        ((Range)exportWorksheet.Cells[1, 10]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 11]).EntireColumn.ColumnWidth = 6;
                        ((Range)exportWorksheet.Cells[1, 12]).EntireColumn.ColumnWidth = 10;
                        ((Range)exportWorksheet.Cells[1, 13]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 14]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 15]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 16]).EntireColumn.ColumnWidth = 20;
                        ((Range)exportWorksheet.Cells[1, 17]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 18]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 19]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 20]).EntireColumn.ColumnWidth = 8;
                        ((Range)exportWorksheet.Cells[1, 21]).EntireColumn.ColumnWidth = 8;


                        Range rgcompany = exportWorksheet.get_Range(exportWorksheet.Cells[1, 1], exportWorksheet.Cells[1, 21]);
                        rgcompany.Application.DisplayAlerts = false;
                        rgcompany.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rgcompany.Font.Bold = true;
                        rgcompany.Font.Size = 15;
                        rgcompany.Merge(0);

                        Range rgtitle = exportWorksheet.get_Range(exportWorksheet.Cells[2, 1], exportWorksheet.Cells[2, 21]);
                        rgtitle.Application.DisplayAlerts = false;
                        rgtitle.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        rgtitle.Font.Bold = true;
                        rgtitle.Font.Size = 15;
                        rgtitle.Merge(0);

                        exportWorksheet.Cells[3, 16] = "单号："+DateTime.Today.ToString("yyyyMMdd")+(new Random()).Next(1,100);
                        ((Range)exportWorksheet.Cells[3, 16]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        ((Range)exportWorksheet.Cells[3, 16]).Font.Bold = true;
                        ((Range)exportWorksheet.Cells[3, 16]).Font.Size = 12;
                        exportWorksheet.Cells[4, 16] = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
                        ((Range)exportWorksheet.Cells[4, 16]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                        ((Range)exportWorksheet.Cells[4, 16]).Font.Bold = true;
                        ((Range)exportWorksheet.Cells[4, 16]).Font.Size = 12;

                        //初始表头值
                        exportWorksheet.Cells[4, 1] = "序号";
                        exportWorksheet.Cells[4, 2] = "流程卡号";
                        exportWorksheet.Cells[4, 3] = "客户代码";
                        exportWorksheet.Cells[4, 4] = "来料型号";
                        exportWorksheet.Cells[4, 5] = "内部型号";
                        //序号	流程卡号	客户代码	来料型号	内部型号

                        exportWorksheet.Cells[4, 6] = "标签型号";
                        exportWorksheet.Cells[4, 7] = "批号";
                        exportWorksheet.Cells[4, 8] = "标签批号";
                        exportWorksheet.Cells[4, 9] = "加工项目";
                        exportWorksheet.Cells[4, 10] = "来料片数";
                        exportWorksheet.Cells[4, 11] = "尺寸";
                        exportWorksheet.Cells[4, 12] = "到货时间";
                        exportWorksheet.Cells[4, 13] = "交期";
                        exportWorksheet.Cells[4, 14] = "测试程序";
                        exportWorksheet.Cells[4, 15] = "PO NO";
                        exportWorksheet.Cells[4, 16] = "备注";
                        exportWorksheet.Cells[4, 17] = "校验码";
                        exportWorksheet.Cells[4, 18] = "工程批";
                        exportWorksheet.Cells[4, 19] = "程序名称";
                        exportWorksheet.Cells[4, 20] = "产品位置";
                        exportWorksheet.Cells[4, 21] = "物料代码";

                        //变工况加粗
                        for (int j = 1; j <= 21; j++)
                        {
                            ((Range)exportWorksheet.Cells[4, j]).Borders.LineStyle = 1;
                        }
                        int total = 0;
                        for (int i = 2; i <= rowsint; i++)
                        {
                            int rowindex = i + 3; //要拷贝到的行
                            int rowindex2 = i ;//源文件的行
                            exportWorksheet.Cells[rowindex, 1]  = i - 1;//序号
                            exportWorksheet.Cells[rowindex, 2] = ((Range)ws.Cells[rowindex2, 6]).Value;//"流程卡号";
                            exportWorksheet.Cells[rowindex, 3] = "";//客户代码
                            exportWorksheet.Cells[rowindex, 4] = ((Range)ws.Cells[rowindex2, 1]).Value;//来料型号
                            exportWorksheet.Cells[rowindex, 5] = "";//内部型号	

                            exportWorksheet.Cells[rowindex, 6] = ws.Cells[rowindex2, 13];//oM[i - 5, 1]; //标签型号
                            exportWorksheet.Cells[rowindex, 7] = ((Range)ws.Cells[rowindex2, 2]).Value;//"批号";
                            exportWorksheet.Cells[rowindex, 8] = ((Range)ws.Cells[rowindex2, 3]).Value;//标签批号
                            exportWorksheet.Cells[rowindex, 9] = ((Range)ws.Cells[rowindex2, 4]).Value;//加工项目
                            exportWorksheet.Cells[rowindex, 10] = ((Range)ws.Cells[rowindex2, 8]).Value;//来料片数	
                            try
                            {
                                total += Convert.ToInt32(((Range)ws.Cells[rowindex2, 8]).Value);
                            }
                            catch
                            { }
                            exportWorksheet.Cells[rowindex, 11] = ((Range)ws.Cells[rowindex2, 12]).Value;//尺寸	
                            exportWorksheet.Cells[rowindex, 12] = ((Range)ws.Cells[rowindex2, 7]).Value; //oG[i - 5, 1].ToString();//到货时间	
                            exportWorksheet.Cells[rowindex, 13] = "";//交期	
                            exportWorksheet.Cells[rowindex, 14] = ((Range)ws.Cells[rowindex2, 15]).Value;//测试程序	
                            exportWorksheet.Cells[rowindex, 15] = "";//PO NO	
                            exportWorksheet.Cells[rowindex, 16] = ((Range)ws.Cells[rowindex2, 10]).Value;//备注	
                            exportWorksheet.Cells[rowindex, 17] = "";//校验码	
                            exportWorksheet.Cells[rowindex, 18] = "";//工程批	
                            exportWorksheet.Cells[rowindex, 19] = ((Range)ws.Cells[rowindex2, 19]).Value;//程序名称	
                            exportWorksheet.Cells[rowindex, 20] = "";//产品位置
                            exportWorksheet.Cells[rowindex, 21] = "";//物料代码

                            for (int m = 1; m <= 21; m++)
                            {
                                ((Range)exportWorksheet.Cells[rowindex, m]).Borders.LineStyle = 1;
                                ((Range)exportWorksheet.Cells[rowindex, m]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignJustify;
                            }

                            //exportWorksheet.Cells[i, 1] = oA[i-5,1];
                            //exportWorksheet.Cells[i, 2] = oB[i - 5, 1];
                        }

                        for (int j = 1; j <= 21; j++)
                        {
                            ((Range)exportWorksheet.Cells[rowsint + 4, j]).Borders.LineStyle = 1;
                            ((Range)exportWorksheet.Cells[rowsint + 5, j]).Borders.LineStyle = 1;
                            ((Range)exportWorksheet.Cells[rowsint + 6, j]).Borders.LineStyle = 1;
                        }

                        //exportWorksheet.Cells[rowsint+4, 9] = "TOTAL";
                        exportWorksheet.Cells[rowsint + 5, 9] = "TOTAL";
                        exportWorksheet.Cells[rowsint + 5, 10] = total;
                        exportWorksheet.Cells[rowsint + 6, 8] = "审核：";
                        exportWorksheet.Cells[rowsint + 6, 16] = "LYSD004A0";

                        exportWorkbook.SaveAs(tbExportPath.Text);
                    }
                    /*
                    for (int i = 5; i <= rowsint; i++)
                    {
                        ws.get_Range("A5", position).
                    }
                    */
                }
                excel.Quit();
                excel = null;
                MessageBox.Show("导出成功！");
                /*
                //取得数据范围区域 (不包括标题列) 
                Range rng1 = ws.Cells.get_Range("B2", "B" + rowsint);   //item


                Range rng2 = ws.Cells.get_Range("K2", "K" + rowsint); //Customer
                object[,] arryItem = (object[,])rng1.Value2;   //get range's value
                object[,] arryCus = (object[,])rng2.Value2;
                //将新值赋给一个数组
                string[,] arry = new string[rowsint - 1, 2];
                for (int i = 1; i <= rowsint - 1; i++)
                {
                    //Item_Code列
                    arry[i - 1, 0] = arryItem[i, 1].ToString();
                    //Customer_Name列
                    arry[i - 1, 1] = arryCus[i, 1].ToString();
                }
                */


            }
            

                /*
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                }
                */

                GC.Collect();


                //写数据

                //写文件
            }
    }
}
