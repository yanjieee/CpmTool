using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Net;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace CpmTool
{



    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            nowAccountNum = 0;

            ThreadPool = new List<Thread>();

            _form2 = new Form2();

            System.Net.ServicePointManager.DefaultConnectionLimit = 50;

            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
        }

        private readonly static int MAX_THREAD = 10;

        private int requireType = 0;
        private List<TAccount> accounts;
        private int nowAccountNum, hasReadedCount;
        private int threadSum = MAX_THREAD;
        private List<Thread> ThreadPool;

        private Form2 _form2;

        private int currentSiteType = 0;
       

        private void button1_Click(object sender, EventArgs e)
        {

            listView1.Columns.Clear();
            listView1.Items.Clear();
            this.currentSiteType = comboBox2.SelectedIndex;
            string company = "";
            if (comboBox3.SelectedIndex == 0)
            {

            } 
            else
            {
                company = comboBox3.SelectedItem.ToString();
            }
            accounts = DB.getInstence().getAccounts(this.currentSiteType, company);
            nowAccountNum = 0;
            hasReadedCount = 0;
            threadSum = MAX_THREAD;
            ListViewItem item = new ListViewItem();
            
            switch (comboBox1.SelectedIndex)
            {
                case 0: item.SubItems.Add("Month To Date"); requireType = 20; break;
                case 1: item.SubItems.Add("Last Month"); requireType = 10; break;
                case 2: item.SubItems.Add("Last 24 Hours"); requireType = 00; break;
                case 3: item.SubItems.Add("Yesterday"); requireType = 30; break;
                default: item.SubItems.Add("Month To Date"); requireType = 20; break;
            }
            item.BackColor = Color.Pink;
            listView1.Items.Add(item);
            this.button1.Enabled = false;
            this.comboBox1.Enabled = false;
            this.comboBox2.Enabled = false;
            ThreadPool.Clear();
            verifyThreadPool();
        }

        private void verifyThreadPool()
        {
            while (threadSum > 0 && nowAccountNum < this.accounts.Count)
            {
                threadSum--;
                TAccount acc = this.accounts[nowAccountNum];
                Network net = new Network(this.currentSiteType, nowAccountNum, requireType, acc.username, acc.password, this.doGetData, this.comboBox4.SelectedIndex);
                Thread thread = new Thread(new ThreadStart(net.run));
                thread.IsBackground = true;
                thread.Start();
                this.ThreadPool.Add(thread);
                nowAccountNum++;
            }
        }

        private void doGetData(List<List<string>> data, bool isSuccess, int dbIndex)
        {
            
            if (this.InvokeRequired)
            {
                this.Invoke(new DelegateDidGetData(doGetData), new object[] { data, isSuccess, dbIndex });
            }
            else
            {
                
                hasReadedCount++;
                if (hasReadedCount >= this.accounts.Count)
                {
                    this.button1.Enabled = true;
                    this.comboBox1.Enabled = true;
                    this.comboBox2.Enabled = true;
                }
                if (isSuccess)
                {
                    if (listView1.Columns.Count == 0)
                    {
                        if (this.currentSiteType == 1)
                        {
                            listView1.Columns.Add("username");
                        }
                        else
                        {
                            listView1.Columns.Add("");
                        }

                        foreach (string s in data[0])
                        {
                            listView1.Columns.Add(s);
                        }

                        listView1.Columns.Add("sitename");
                        listView1.Columns.Add("company");
                        listView1.Columns.Add("volume");
                        listView1.Columns.Add("revenue");

                        foreach (ColumnHeader h in listView1.Columns)
                        {
                            h.Width = -2;
                        }
                    }
                    ListViewItem item;
                    if (this.currentSiteType == 1)
                    {
                        item = new ListViewItem(accounts[dbIndex].username);
                    }
                    else
                    {
                        item = new ListViewItem();
                    }

                    foreach (string s in data[1])
                    {
                        item.SubItems.Add(s);
                    }

                    TAccount acc = DB.getInstence().getAccount(accounts[dbIndex].ID);

                    item.SubItems.Add(acc.sitename);
                    item.SubItems.Add(acc.company);
                    item.SubItems.Add(acc.volume.ToString());
                    item.SubItems.Add(acc.revenue.ToString());

                    if (accounts[dbIndex].important)
                    {
                        item.ForeColor = Color.Red;
                    }

                    item.Tag = dbIndex;
                    listView1.Items.Add(item);
                }
                else
                {
                    //获取失败
                    ListViewItem item = new ListViewItem();
                    item.SubItems.Add(accounts[dbIndex].username);
                    item.SubItems.Add("无数据");
                    item.ForeColor = Color.Gray;
                    item.Tag = dbIndex;
                    listView1.Items.Add(item);
                }
                

                
                threadSum++;
                verifyThreadPool();

            }
            
        }

        public void ExportExcel(ListView lv)
        {
            if (lv.Items == null) return;

            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = DateTime.Now.ToString("yyyy-MM-dd");
            saveDialog.ShowDialog();
            saveFileName = saveDialog.FileName;
            if (saveFileName.IndexOf(":") < 0)
                return;
            //这里直接删除，因为saveDialog已经做了文件是否存在的判断
            try
            {
                if (File.Exists(saveFileName)) File.Delete(saveFileName);
            }
            catch (Exception e1)
            {
                MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + e1.Message);
                return;
            }
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("无法创建Excel对象，可能您的机器未安装Excel");
                return;
            }
            this.Enabled = false;
            Excel.Workbooks workbooks = xlApp.Workbooks;
            Excel.Workbook workbook = workbooks.Add(true);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            xlApp.Visible = false;
            //填充列
            for (int i = 0; i < lv.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = lv.Columns[i].Text.ToString();
                ((Excel.Range)worksheet.Cells[1, i + 1]).Font.Bold = true;
            }

            //填充数据（这里分了两种情况，1：lv带CheckedBox，2：不带CheckedBox）

            //带CheckedBoxes
            if (lv.CheckBoxes == true)
            {
                int tmpCnt = 0;
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    if (lv.Items[i].Checked == true)
                    {
                        for (int j = 0; j < lv.Columns.Count; j++)
                        {
                            if (j == 0)
                            {
                                worksheet.Cells[2 + tmpCnt, j + 1] = lv.Items[i].Text.ToString();
                                ((Excel.Range)worksheet.Cells[2 + tmpCnt, j + 1]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            }
                            else
                            {
                                worksheet.Cells[2 + tmpCnt, j + 1] = lv.Items[i].SubItems[j].Text.ToString();
                                ((Excel.Range)worksheet.Cells[2 + tmpCnt, j + 1]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            }
                        }
                        tmpCnt++;

                        Application.DoEvents();
                    }
                }
            }
            else  //不带Checkedboxe
            {
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    for (int j = 0; j < lv.Items[i].SubItems.Count; j++)
                    {
                        if (j == 0)
                        {
                            worksheet.Cells[2 + i, j + 1] = lv.Items[i].Text.ToString();
                            ((Excel.Range)worksheet.Cells[2 + i, j + 1]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                        else
                        {
                            worksheet.Cells[2 + i, j + 1] = lv.Items[i].SubItems[j].Text.ToString();
                            ((Excel.Range)worksheet.Cells[2 + i, j + 1]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                    }
                }
            }

            for (int j = 0; j < 6; j++)
            {
                if (j == 3)
                {
                    ((Excel.Range)worksheet.Cells[1, j + 1]).EntireColumn.ColumnWidth = 40;
                    continue;
                }
                ((Excel.Range)worksheet.Cells[1, j + 1]).EntireColumn.AutoFit();
            }

            if (lv.CheckBoxes == true)
            {
                ((Excel.Range)worksheet.Cells[1, 1]).EntireColumn.Delete(Excel.XlDirection.xlDown);
            }

            object missing = System.Reflection.Missing.Value;
            try
            {
                workbook.Saved = true;
                //Excel.XlFileFormat.xlExcel8 不支持office 2007以下版本
                workbook.SaveAs(saveFileName, Excel.XlFileFormat.xlExcel8, missing, missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
            }
            catch (Exception e1)
            {
                MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + e1.Message);
            }
            finally
            {
                xlApp.Quit();
                System.GC.Collect();
            }
            this.Enabled = true;
            MessageBox.Show("导出成功！");
        }


        private void button2_Click(object sender, EventArgs e)
        {
            this.button1.Enabled = true;
            this.comboBox1.Enabled = true;
            foreach(Thread th in this.ThreadPool)
            {
                if (th.IsAlive)
                {
                    th.Abort();
                }
            }
            this.ThreadPool.Clear();
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }
            ListViewItem itemClicked = listView1.GetItemAt(e.X, e.Y);
            if (itemClicked != null && itemClicked.Tag != null) 
            {
                _form2.StartPosition = FormStartPosition.CenterParent;
                _form2.Tag = itemClicked.Tag;
                _form2.sitetype = this.currentSiteType;
                _form2.timezone = this.comboBox4.SelectedIndex;
                _form2.ShowDialog();
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(comboBox2.SelectedIndex)
            {
                case 0: 
                    comboBox1.Items[2] = "Last 24 Hours";
                    if (comboBox1.Items.Count > 3)
                    {
                        comboBox1.Items.RemoveAt(3);
                    }
                    break;
                case 1: 
                    comboBox1.Items[2] = "Last 48 Hours"; 
                    comboBox1.Items.Add("Yesterday"); 
                    break;
            }
            comboBox3.Items.Clear();
            comboBox3.Items.Add("全部");
            List<string> companyList = DB.getInstence().getCompanys(comboBox2.SelectedIndex);
            foreach (string s in companyList)
            {
                comboBox3.Items.Add(s);
            }
            comboBox3.SelectedIndex = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //this.button3.Enabled = false;
            if (this.listView1.Items.Count > 0)
            {
                this.ExportExcel(listView1);
            }
            else
            {
                MessageBox.Show("没有数据");
            }
            
        }
    }
}
