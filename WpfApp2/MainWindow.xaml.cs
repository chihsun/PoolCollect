using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
//using Excel = Microsoft.Office.Interop.Excel;
using SpreadsheetLight;

namespace WpfApp2
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
        }
        public List<PoolData> PDatas = new List<PoolData>();
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            List<string> myList = new List<string>();
            string folderName = System.Environment.CurrentDirectory + @"\Data";
            int count = 0;
            foreach (var finame in System.IO.Directory.GetFileSystemEntries(folderName))
            {
                //MessageBox.Show(System.IO.Path.GetExtension(finame));
                if (System.IO.Path.GetExtension(finame) != ".xlsx")
                    continue;
                myList.Add(System.IO.Path.GetFileNameWithoutExtension(finame));
                GetPoolData(finame);
                this.LB_1.Content = count;
                count++;
            }
            var sortdata = PDatas;
            sortdata.Sort((x, y) => { return x.Room.CompareTo(y.Room); });
            /*
            foreach (var x in PDatas)
            {
                this.TXTB1.Text += x.ToString() + Environment.NewLine;
            }
            foreach (var x in myList)
                this.LB_1.Content += x + Environment.NewLine;
            */
            /*
            
            Excel.Application Excel_APP1 = new Excel.Application();
            Excel.Workbook Excel_WB1 = Excel_APP1.Workbooks.Add();
            Excel.Worksheet Excel_WS1 = Excel_WB1.Worksheets[1];
            Excel_WS1.Cells[1, 1] = "科室";
            Excel_WS1.Cells[1, 2] = "姓名";
            Excel_WS1.Cells[1, 3] = "代號";
            Excel_WS1.Cells[1, 4] = "計數";
            Excel_WS1.Cells[1, 5] = "備註說明";

            try
            {
                for (int i = 0; i < PDatas.Count; i++)
                {
                    Excel_WS1.Cells[i + 2, 1] = PDatas[i].Room;
                    Excel_WS1.Cells[i + 2, 2] = PDatas[i].Name;
                    Excel_WS1.Cells[i + 2, 3] = PDatas[i].ID;
                    Excel_WS1.Cells[i + 2, 4] = PDatas[i].Points;
                    Excel_WS1.Cells[i + 2, 5] = PDatas[i].Detial;
                    this.TXTB1.Text += PDatas[i].ToString() + Environment.NewLine;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            */
            SLDocument sl = new SLDocument();
            sl.SetCellValue(1, 1, "科室");
            sl.SetCellValue(1, 2, "姓名");
            sl.SetCellValue(1, 3, "代號");
            sl.SetCellValue(1, 4, "計數");
            sl.SetCellValue(1, 5, "備註說明");
            sl.SetColumnWidth(1, 20);
            sl.SetColumnWidth(2, 15);
            sl.SetColumnWidth(3, 10);
            sl.SetColumnWidth(4, 10);
            sl.SetColumnWidth(5, 100);
            SLStyle style = sl.CreateStyle();
            style.Alignment.WrapText = true;
            style.Alignment.Vertical = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center;
            style.Alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
            for (int i = 0; i < 5; i++)
            {
                if (i == 4)
                    style.Alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left;
                sl.SetColumnStyle(i + 1, style);
            }
            try
            {
                for (int i = 0; i < PDatas.Count; i++)
                {
                     sl.SetCellValue(i + 2, 1, PDatas[i].Room);
                    sl.SetCellValue(i + 2, 2, PDatas[i].Name);
                    sl.SetCellValue(i + 2, 3, int.TryParse(PDatas[i].ID, out int id) ? id : 0);
                    sl.SetCellValue(i + 2, 4, PDatas[i].Points);
                    sl.SetCellValue(i + 2, 5, PDatas[i].describe);
                    style.Fill.SetPattern(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid, (i + 2) % 2 == 0 ? System.Drawing.ColorTranslator.FromHtml("#c2f0c2") : System.Drawing.ColorTranslator.FromHtml("#eafaea"), System.Drawing.Color.Black);
                    for (int j = 0; j < 5; j++)
                        sl.SetCellStyle(i + 2, j + 1, style);
                    this.TXTB1.Text += PDatas[i].ToString() + Environment.NewLine;
                }
                sl.SaveAs(System.Environment.CurrentDirectory + @"\result-" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx");
                MessageBox.Show("資料匯出成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //if (System.IO.File.Exists("result.txt"))
            //    System.IO.File.Delete("result.txt");
            //if (System.IO.File.Exists(Refile))
            //    System.IO.File.Delete(Refile);

            //System.IO.File.WriteAllText("result.txt", this.TXTB1.Text);
            //sl.SaveAs(Refile);
            /*
            Excel_WB1.SaveAs(Refile);
            Excel_WB1.Close();
            Excel_APP1.Quit();
            */
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            this.LB_1.Content = string.Empty;
        }

        public void GetPoolData(string fname)
        {
            /*
            string FileStr = @"人體試驗(IRB)委員會";
            string FilePat = System.Environment.CurrentDirectory;
            string FileName = FilePat + @"\" + FileStr + ".xlsx";
            if (System.IO.File.Exists(FilePat + @"\" + FileStr + ".xlsx"))
                this.LB_1.Content = FileStr;
            else
                return;
            */
            
            if (!System.IO.File.Exists(fname))
                return;
            SLDocument sl = new SLDocument(fname);
            try
            {
                for (int i = 0; i < 50; i++)
                {
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 3, 1)))
                        break;
                    if (string.IsNullOrEmpty(sl.GetCellValueAsString(i + 3, 2))
                        || string.IsNullOrEmpty(sl.GetCellValueAsString(i + 3, 3))
                        || !Double.TryParse(sl.GetCellValueAsString(i + 3, 13).Trim(), out double points)
                        || points == 0)
                    {
                        this.TXTB1.Text += "錯誤: "
                            + System.IO.Path.GetFileNameWithoutExtension(fname) + sl.GetCellValueAsString(i + 3, 2).Trim()
                        + Environment.NewLine;
                        continue;
                    }
                    PoolData data = new PoolData();
                    data.Room = sl.GetCellValueAsString(i + 3, 1).Trim();
                    data.Name = sl.GetCellValueAsString(i + 3, 2).Trim();
                    data.ID = sl.GetCellValueAsString(i + 3, 3).Trim();
                    data.Points = points;
                    data.Detail.Add(System.IO.Path.GetFileNameWithoutExtension(fname) + " " + data.Points.ToString());

                    if (PDatas.Count > 0)
                    {
                        bool duplicated = false;
                        foreach (var x in PDatas)
                        {
                            if (x.ID == data.ID && x.Name == data.Name)
                            {
                                x.Points += data.Points;
                                x.Detail.Add(data.Detail.FirstOrDefault());
                                duplicated = true;
                                break;
                            }
                        }
                        if (!duplicated)
                            PDatas.Add(data);
                    }
                    else
                        PDatas.Add(data);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            sl.CloseWithoutSaving();
            /*
            Excel.Application Excel_APP1 = new Excel.Application();
            Excel.Workbook Excel_WB1 = Excel_APP1.Workbooks.Open(fname);
            Excel.Worksheet Excel_WS1 = Excel_WB1.Worksheets[1];
            try
            {
                for (int i = 0; i < 20; i++)
                {
                    if (string.IsNullOrEmpty(Excel_WS1.Cells[i + 3, 1].Value))
                        break;
                    PoolData data = new PoolData();
                    data.Room = Excel_WS1.Cells[i + 3, 1].Value.ToString().Trim();
                    data.Name = Excel_WS1.Cells[i + 3, 2].Value.ToString().Trim();
                    data.ID = Excel_WS1.Cells[i + 3, 3].Value.ToString().Trim();
                    data.Points = Convert.ToDouble(Excel_WS1.Cells[i + 3, 13].Value.ToString());
                    data.Detial = System.IO.Path.GetFileNameWithoutExtension(fname) + " " + data.Points.ToString() + ";";

                    if (PDatas.Count > 0)
                    {
                        bool duplicated = false;
                        foreach (var x in PDatas)
                        {
                            if (x.ID == data.ID && x.Name == data.Name)
                            {
                                x.Points += data.Points;
                                x.Detial += data.Detial;
                                duplicated = true;
                                break;
                            }
                        }
                        if (!duplicated)
                            PDatas.Add(data);
                    }
                    else
                        PDatas.Add(data);
                }
                /*
                foreach (var x in PDatas)
                {
                    this.TXTB1.Text += x.ToString() + Environment.NewLine;
                }
                */
            /*
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
        Excel_WB1.Close();
        Excel_APP1.Quit();
        */
        }
        public class PoolData
        {
            public string Room { get; set; }
            public string Name { get; set; }
            public string ID { get; set; }
            public double Points { get; set; }
            public bool Director { get; set; }
            //public string Detail { get; set; }
            public List<string> Detail;
            public string describe
            {
                get
                {
                    return string.Join(";", Detail);
                }
            }
            public override string ToString()
            {
                return Room + "," + Name + "," + ID + "," + Points.ToString() + "," + string.Join(";", Detail);
            }
            public PoolData()
            {
                Detail = new List<string>();
            }

        }
    }
}
