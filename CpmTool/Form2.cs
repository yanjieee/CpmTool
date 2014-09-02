using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace CpmTool
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private Thread myThread;
        public int sitetype = 0;
        public int timezone = 0;

        private void Form2_Shown(object sender, EventArgs e)
        {
            TAccount acc = DB.getInstence().getAccounts(this.sitetype)[(int)this.Tag];
            listView1.Items.Clear();
            listView1.Columns.Clear();

            this.Text = "正在读取...";

            Network net = new Network(this.sitetype, (int)this.Tag, 21, acc.username, acc.password, this.doGetData, this.timezone);
            myThread = new Thread(new ThreadStart(net.run));
            myThread.IsBackground = true;
            myThread.Start();
        }


        private void doGetData(List<List<string>> data, bool isSuccess, int dbIndex)
        {

            if (this.InvokeRequired)
            {
                this.Invoke(new DelegateDidGetData(doGetData), new object[] { data, isSuccess, dbIndex });
            }
            else
            {
                this.Text = "";
                listView1.Items.Clear();

                if (listView1.Columns.Count == 0)
                {
                    listView1.Columns.Add("");
                    foreach (string s in data[0])
                    {
                        listView1.Columns.Add(s);
                    }
                    foreach (ColumnHeader h in listView1.Columns)
                    {
                        h.Width = -2;
                    }
                }

                for (int i = 1; i < data.Count;i++ )
                {
                    ListViewItem item = new ListViewItem();
                    foreach (string s in data[i])
                    {
                        item.SubItems.Add(s);
                    }
                    listView1.Items.Add(item);
                }
            }

        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (myThread != null && myThread.IsAlive)
            {
                myThread.Abort();
            }
        }
    }
}
