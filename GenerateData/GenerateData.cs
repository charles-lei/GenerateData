using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace GenerateData
{
    public partial class GenerateData : Form
    {
        public GenerateData()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = null;
            openFileDialog1.Filter = "EXCLE(*.XLSX)|*.XLSX";
            if (openFileDialog1.ShowDialog() == DialogResult.OK && openFileDialog1.OpenFile() != null)
            {
                txtPath.Text = openFileDialog1.FileName;

            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            XSSFWorkbook wk;
            using (FileStream fs = File.OpenRead(txtPath.Text))   //打开myxls.xls文件
            {
                 wk = new XSSFWorkbook(fs);   //把xls文件中的数据写入wk中
               
            }
            ISheet sheetLead = wk.GetSheet("Lead");
            if (sheetLead == null)
            {
                MessageBox.Show("Cannot find 'Lead' sheet");
                return;
            }
            ISheet sheetClosedOpp = wk.GetSheet("Closed Opp");
            if (sheetClosedOpp == null)
            {
                MessageBox.Show("Cannot find 'Closed Opp' sheet");
                return;
            }
            ISheet sheetPipeline = wk.GetSheet("Pipeline");
            if (sheetPipeline == null)
            {
                MessageBox.Show("Cannot find 'Pipeline' sheet");
                return;
            }
            var leadData = GetSheetData(sheetLead);
            var closedOppData = GetSheetData(sheetClosedOpp);
            var pipelineData = GetSheetData(sheetPipeline);
            var verticals = new List<string>();
            leadData.ForEach(l => {
                if (!verticals.Contains(l[0]))
                {
                    verticals.Add(l[0]);
                }
            });
            closedOppData.ForEach(l => {
                if (!verticals.Contains(l[0]))
                {
                    verticals.Add(l[0]);
                }
            });
            pipelineData.ForEach(l => {
                if (!verticals.Contains(l[0]))
                {
                    verticals.Add(l[0]);
                }
            });
           
            verticals.Add("Grand Total");
            //创建工作薄
            XSSFWorkbook wkNew = new XSSFWorkbook();
            List<string> leadStatus = new List<string>{ "Leads", "New","Working","Accepted","Converted","Rejected" };
            Dictionary<string,string> pipelineStatus = new Dictionary<string, string> { { "ClosedWon", "Closed Won" }, { "ClosedLost", "Closed Lost" }, { "ClosedWonM", "Closed Won" }, { "ClosedLostM", "Closed Lost" }, { "PipelineCreated", "" } };
            //创建一个名称为mySheet的表
            ISheet tb = wkNew.CreateSheet("ReferenceData");
            for (int i = 0; i< verticals.Count;i++)
            {
                //创建一行，此行为第二行
                for (int j = 0; j < leadStatus.Count; j++)
                {
                    IRow row = tb.CreateRow(i* (leadStatus.Count+pipelineStatus.Count)+j);
                    ICell cell = row.CreateCell(0);  //在第二行中创建单元格
                    cell.SetCellValue(verticals[i]);//循环往第二行的单元格中添加数据

                    ICell cell1 = row.CreateCell(1);
                    cell1.SetCellValue(leadStatus[j]);

                    for (int k = 2; k < 12 + 2; k++)
                    {
                        int data = 0;
                        ICell cellData = row.CreateCell(k);
                        if (verticals[i] == "Grand Total")
                        {
                            if (j == 0)
                                data = leadData.Where(L => L[1].StartsWith((k - 1).ToString() + "/")).Count();
                            else
                                data = leadData.Where(L => L[1].StartsWith((k - 1).ToString() + "/") && L[2] == leadStatus[j]).Count();
                            cellData.SetCellValue(data);
                        }
                        else {
                            if (j == 0)
                                data = leadData.Where(L => L[0] == verticals[i] && L[1].StartsWith((k - 1).ToString() + "/")).Count();
                            else
                                data = leadData.Where(L => L[0] == verticals[i] && L[1].StartsWith((k - 1).ToString() + "/") && L[2] == leadStatus[j]).Count();
                            cellData.SetCellValue(data);
                        }
                    }
                }
                for (int j = 0; j < pipelineStatus.Count;j++)
                {
                    IRow row = tb.CreateRow(i * (leadStatus.Count + pipelineStatus.Count)+leadStatus.Count + j);
                    ICell cell = row.CreateCell(0);  //在第二行中创建单元格
                    cell.SetCellValue(verticals[i]);//循环往第二行的单元格中添加数据

                    ICell cell1 = row.CreateCell(1);
                    cell1.SetCellValue(pipelineStatus.ElementAt(j).Key);

                    for (int k = 2; k < 12 + 2; k++)
                    {
                        ICell cellData = row.CreateCell(k);
                        if (j == 0 || j == 1)
                        {
                            if(verticals[i]=="Grand Total")
                            {
                                int data = closedOppData.Where(L=> L[1].StartsWith((k - 1).ToString() + "/") && L[2] == pipelineStatus.ElementAt(j).Value).Count();
                                cellData.SetCellValue(data);
                            }
                            else
                            {
                                int data = closedOppData.Where(L => L[0] == verticals[i] && L[1].StartsWith((k - 1).ToString() + "/") && L[2] == pipelineStatus.ElementAt(j).Value).Count();
                                cellData.SetCellValue(data);
                            }
                        }
                        else if (j == 2 || j == 3)
                        {
                            if (verticals[i] == "Grand Total")
                            {
                                double data = closedOppData.Where(L => L[1].StartsWith((k - 1).ToString() + "/") && L[2] == pipelineStatus.ElementAt(j).Value).Sum(L =>
                                {
                                    double ret = 0.0;
                                    double.TryParse(L[4], out ret);
                                    return ret;
                                });
                                cellData.SetCellValue(data);
                            }
                            else
                            {
                                double data = closedOppData.Where(L => L[0] == verticals[i] && L[1].StartsWith((k - 1).ToString() + "/") && L[2] == pipelineStatus.ElementAt(j).Value).Sum(L =>
                                {
                                    double ret = 0.0;
                                    double.TryParse(L[4], out ret);
                                    return ret;
                                });
                                cellData.SetCellValue(data);
                            }
                        }
                        else
                        {
                            if (verticals[i] == "Grand Total")
                            {
                                double data = pipelineData.Where(L => L[1].StartsWith((k - 1).ToString() + "/")).Sum(L => { double ret = 0.0; double.TryParse(L[4], out ret); return ret; });
                                cellData.SetCellValue(data);
                            }
                            else
                            {
                                double data = pipelineData.Where(L => L[0] == verticals[i] && L[1].StartsWith((k - 1).ToString() + "/")).Sum(L => { double ret = 0.0; double.TryParse(L[4], out ret); return ret; });
                                cellData.SetCellValue(data);
                            }
                        }
                    }
                }
               
            }


            using (FileStream fs1 = File.OpenWrite(Application.StartupPath+@"/myxls.xlsx")) //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建是不要打开该文件！
            {
                wkNew.Write(fs1);   //向打开的这个xls文件中写入mySheet表并保存。
                MessageBox.Show("提示：创建成功！");
            }
        }

        private List<List<string>> GetSheetData(ISheet sheet)
        {
            var rows = new List<List<string>>();
            if (sheet == null)
                return rows;
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    var leadRow = new List<string>();
                    for (int j = 0; j <= row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            leadRow.Add(cell.ToString());
                        }
                    }
                    if(leadRow.Count>=4 && !leadRow.All(s=>string.IsNullOrEmpty(s)))
                        rows.Add(leadRow);
                }
            }
            return rows;
        }
    }
}
