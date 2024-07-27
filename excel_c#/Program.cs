using System;
using OfficeOpenXml;
using System.IO;
using System.Windows.Forms;


namespace excel_c_
{
    internal class Program
    {

        public static string summaryFilePath;
        public static string ddicFilePath;
        public static string header_initiative;
        public static string cellValue;
        public static double doubleValue;
        public static int col2;

        [STAThread]
        static void Main(string[] args)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择总表的Excel文件";
            openFileDialog.Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            summaryFilePath = booleanPathed(openFileDialog);

            if (summaryFilePath.Equals("err"))
            {
                MessageBox.Show("操作错误！", "请从新打开软件");
                Environment.Exit(0);
            }

            OpenFileDialog openFileDialogDDIC = new OpenFileDialog();
            openFileDialogDDIC.Title = "选择字典表的Excel文件";
            openFileDialogDDIC.Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            ddicFilePath = booleanPathed(openFileDialogDDIC);

            if (ddicFilePath.Equals("err"))
            {
                MessageBox.Show("操作错误！", "请从新打开软件");
                Environment.Exit(0);
            }

            MatchingAssignment(summaryFilePath, ddicFilePath);

            MessageBox.Show("完成写入！", "程序运行结束");
            Environment.Exit(0);
        }

        static string booleanPathed(OpenFileDialog openFileDialog)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string FilePath = openFileDialog.FileName;
                if (!File.Exists(FilePath))
                {
                    MessageBox.Show("文件不存在，请检查路径！", "err");
                    Environment.Exit(0);
                }
                //Console.WriteLine("操作完成。");
                return (FilePath);
            }
            return ("err");
        }


        static void MatchingAssignment(string summaryFilePath, string ddicFilePath)
        {
            string ID = "";
            // 使用EPPlus打开Excel文件  
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 设置许可证上下文为非商业  
            using (var package = new ExcelPackage(new FileInfo(summaryFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // 假设我们要读取第一行第一列的值  
                if (worksheet.Dimension.Address != null)
                {
                    col2 = worksheet.Dimension.End.Column;
                    // 你也可以遍历整个工作表  
                    for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                        {
                            //当前单元格标头
                            header_initiative = worksheet.Cells[1, col].Value?.ToString();

                            cellValue = worksheet.Cells[row, col].Value?.ToString();

                            ID = LookUpDdic(ddicFilePath, header_initiative, cellValue);
                            //Console.Write($"单元格的值是: {cellValue}");

                            //写进去
                            for (int j = 1; j <= col2; j++)
                            {
                                if (ID.Equals(worksheet.Cells[1, j].Value?.ToString()))//找列
                                {
                                    if (Convert.ToInt32(worksheet.Cells[row, j].Value) != 1)//判断它之前是否为1
                                    {
                                        //if (!string.IsNullOrEmpty(cellValue))//********不等于空就写入1********
                                        if (!string.IsNullOrEmpty(cellValue) && (double.TryParse(cellValue, out doubleValue) && doubleValue != 0))//不等于空并且不等于0就写入1
                                        {
                                            worksheet.Cells[row, j].Value = 1;
                                            break;
                                        }
                                        else
                                        {
                                            worksheet.Cells[row, j].Value = 0;
                                            break;
                                        }
                                    }
                                }

                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("工作表是空的。");
                }
                package.Save();
            }
        }

        static string LookUpDdic(string ddicFilePath, string header_initiative, string cellValue)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 设置许可证上下文为非商业  
            using (var package = new ExcelPackage(new FileInfo(ddicFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension.Address != null)
                {
                    for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                    {
                        //只需要遍历第二列就行
                        var cellValueDdic = worksheet.Cells[row, 2].Value?.ToString();

                        if (header_initiative.Equals(cellValueDdic))
                        {
                            return worksheet.Cells[row, 2 + 1].Value?.ToString();
                        }
                    }
                    return "nohaveID";
                }
            }
            return "nohaveID";
        }

    }
}

