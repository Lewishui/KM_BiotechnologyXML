using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using clsdatabaseinfo;
using clsKMBuiness;
using clsKMCommon;
using WeifenLuo.WinFormsUI.Docking;

namespace KM_BiotechnologyXML
{
    public partial class frmDataCenter : DockContent
    {

        List<xmlDataSources> Results;
        private SortableBindingList<xmlDataSources> sortablePendingOrderList;
        int RowRemark = 0;
        int cloumn = 0;
        bool checkedav = false;
        public log4net.ILog ProcessLogger;
        public log4net.ILog ExceptionLogger;
        public frmDataCenter()
        {
            InitializeComponent();
            InitialSystemInfo();

            ProcessLogger.Fatal("07932:System Login Start " + DateTime.Now.ToString());

            clsAllnew BusinessHelp = new clsAllnew();

            List<xmlDataSources> StatusResults = BusinessHelp.findALLStatus_Server();
            ProcessLogger.Fatal("07935:System Login successful mapping data Start " + DateTime.Now.ToString());

            var nations = StatusResults.Select(o => new MockEntity { ShortName = o.Rack_ID, FullName = o.Rack_ID }).ToList();
            nations.Insert(0, new MockEntity { ShortName = "", FullName = "所有" });
            this.comboBox1.DisplayMember = "FullName";
            this.comboBox1.ValueMember = "ShortName";
            this.comboBox1.DataSource = nations;

            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync(new WorkerArgument { OrderCount = 0, CurrentIndex = 0 });
            }

        }
        private void InitialSystemInfo()
        {
            #region 初始化配置
            ProcessLogger = log4net.LogManager.GetLogger("ProcessLogger");
            ExceptionLogger = log4net.LogManager.GetLogger("SystemExceptionLogger");
            ProcessLogger.Fatal("System Start " + DateTime.Now.ToString());
            #endregion
        }
        private void btfind_Click(object sender, EventArgs e)
        {
            InitializeDataSource();
        }
        private void InitializeDataSource()
        {
            Results = new List<xmlDataSources>();

            clsAllnew BusinessHelp = new clsAllnew();
            string start_time = clsCommHelp.objToDateTime1(dateTimePicker1.Text);
            string end_time = clsCommHelp.objToDateTime1(dateTimePicker2.Text);
            string Rack_ID = this.comboBox1.Text;
            if (Rack_ID == "所有")
                Rack_ID = "";

            Results = BusinessHelp.findXML_Server(this.keywordTextBox.Text, start_time, end_time, Rack_ID);

            //if (Result)
            var q = (from o in Results
                     orderby o.Rack_ID descending
                     select o);
            this.dataGridView1.AutoGenerateColumns = false;
            sortablePendingOrderList = new SortableBindingList<xmlDataSources>(q.ToList());
            this.bindingSource1.DataSource = sortablePendingOrderList;
            //  this.dataGridView1.DataSource = null;
            //  this.bindingSource1.Sort = "Rack_ID  ASC";
            this.dataGridView1.DataSource = this.bindingSource1;

            toolStripLabel2.Text = "条目：" + Results.Count.ToString();

        }
        public class SortableBindingList<T> : BindingList<T>
        {
            private bool isSortedCore = true;
            private ListSortDirection sortDirectionCore = ListSortDirection.Ascending;
            private PropertyDescriptor sortPropertyCore = null;
            private string defaultSortItem;

            public SortableBindingList() : base() { }

            public SortableBindingList(IList<T> list) : base(list) { }

            protected override bool SupportsSortingCore
            {
                get { return true; }
            }

            protected override bool SupportsSearchingCore
            {
                get { return true; }
            }

            protected override bool IsSortedCore
            {
                get { return isSortedCore; }
            }

            protected override ListSortDirection SortDirectionCore
            {
                get { return sortDirectionCore; }
            }

            protected override PropertyDescriptor SortPropertyCore
            {
                get { return sortPropertyCore; }
            }

            protected override int FindCore(PropertyDescriptor prop, object key)
            {
                for (int i = 0; i < this.Count; i++)
                {
                    if (Equals(prop.GetValue(this[i]), key)) return i;
                }
                return -1;
            }

            protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
            {
                isSortedCore = true;
                sortPropertyCore = prop;
                sortDirectionCore = direction;
                Sort();
            }

            protected override void RemoveSortCore()
            {
                if (isSortedCore)
                {
                    isSortedCore = false;
                    sortPropertyCore = null;
                    sortDirectionCore = ListSortDirection.Ascending;
                    Sort();
                }
            }

            public string DefaultSortItem
            {
                get { return defaultSortItem; }
                set
                {
                    if (defaultSortItem != value)
                    {
                        defaultSortItem = value;
                        Sort();
                    }
                }
            }

            private void Sort()
            {
                List<T> list = (this.Items as List<T>);
                list.Sort(CompareCore);
                ResetBindings();
            }

            private int CompareCore(T o1, T o2)
            {
                int ret = 0;
                if (SortPropertyCore != null)
                {
                    ret = CompareValue(SortPropertyCore.GetValue(o1), SortPropertyCore.GetValue(o2), SortPropertyCore.PropertyType);
                }
                if (ret == 0 && DefaultSortItem != null)
                {
                    PropertyInfo property = typeof(T).GetProperty(DefaultSortItem, BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.IgnoreCase, null, null, new Type[0], null);
                    if (property != null)
                    {
                        ret = CompareValue(property.GetValue(o1, null), property.GetValue(o2, null), property.PropertyType);
                    }
                }
                if (SortDirectionCore == ListSortDirection.Descending) ret = -ret;
                return ret;
            }

            private static int CompareValue(object o1, object o2, Type type)
            {
                if (o1 == null) return o2 == null ? 0 : -1;
                else if (o2 == null) return 1;
                else if (type.IsPrimitive || type.IsEnum) return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
                else if (type == typeof(DateTime)) return Convert.ToDateTime(o1).CompareTo(o2);
                else return String.Compare(o1.ToString().Trim(), o2.ToString().Trim());
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.RowIndex < 0)
            //    return;
            //if (RowRemark >= dataGridView1.RowCount || RowRemark == -1)
            //    return;

            //var row = dataGridView1.Rows[e.RowIndex];
            //var model = row.DataBoundItem as xmlDataSources;
            //List<xmlDataSources> updateResults = new List<xmlDataSources>();
            //if (model.Id != null && model.Id != "")
            //    updateResults.Add(model);

            //clsAllnew BusinessHelp = new clsAllnew();

            //BusinessHelp.update_OrderServer(updateResults);
            //InitializeDataSource();

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            RowRemark = e.RowIndex;
            cloumn = e.ColumnIndex;
            if (RowRemark == -1)
                return;

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex == 0)
                return;
            if (RowRemark >= dataGridView1.RowCount || RowRemark == -1)
                return;

            var row = dataGridView1.Rows[e.RowIndex];
            var model = row.DataBoundItem as xmlDataSources;
            List<xmlDataSources> updateResults = new List<xmlDataSources>();
            if (model.Id != null && model.Id != "")
                updateResults.Add(model);

            clsAllnew BusinessHelp = new clsAllnew();

            BusinessHelp.update_OrderServer(updateResults);
            InitializeDataSource();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            List<xmlDataSources> FilterOrderResults = new List<xmlDataSources>();

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if ((bool)dataGridView1.Rows[i].Cells[0].EditedFormattedValue == true)
                {

                    var row = dataGridView1.Rows[i];
                    var model = row.DataBoundItem as xmlDataSources;
                    FilterOrderResults.Add(model);
                }
            }
            if (FilterOrderResults.Count > 0)
            {
                xmldown(FilterOrderResults);
                MessageBox.Show("下载完成,请查看", "下载", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
                MessageBox.Show("您还未选中数据", "下载", MessageBoxButtons.OK, MessageBoxIcon.Error);

            // NewMethod();

        }

        private void xmldown(List<xmlDataSources> FilterOrderResults)
        {
            #region 获取模板路径
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            string fullPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\"), "T.xml");
            SaveFileDialog sfdDownFile = new SaveFileDialog();
            sfdDownFile.OverwritePrompt = false;
            string DesktopPath = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            sfdDownFile.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");
            //Retrieve-20170512-114950.order
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            file = "";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                file = dialog.SelectedPath;
                //MessageBox.Show("已选择文件夹:" + foldPath, "选择文件夹提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            sfdDownFile.FileName = Path.Combine(file, "Retrieve-" + DateTime.Now.ToString("yyyyMMdd"));

            string strExcelFileName = string.Empty;

            #endregion
            #region 导出前校验模板信息
            if (string.IsNullOrEmpty(sfdDownFile.FileName))
            {
                MessageBox.Show("File name can't be empty, please confirm, thanks!");
                return;
            }
            if (!File.Exists(fullPath))
            {
                MessageBox.Show("Template file does not exist, please confirm, thanks!");
                return;
            }
            else
            {
                strExcelFileName = sfdDownFile.FileName + ".xlsx";
            }
            if (file == "")
            {
                return;
            }
            #endregion
            foreach (xmlDataSources item in FilterOrderResults)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(fullPath);
                XmlElement xmlRoot = xmlDoc.DocumentElement;
                //xmlDoc.InnerXml.Replace("A1", "8888");

                XmlNodeList personNodes = xmlRoot.GetElementsByTagName("Hole"); //获取Person子节点集合
                foreach (XmlNode node in personNodes)
                {
                    //Hole_ID
                    XmlElement ele = (XmlElement)node;
                    if (ele.GetAttribute("ID") == "A1")
                    {
                        ele.SetAttribute("ID", item.Hole_ID);

                    }
                    //Tube_ID
                    foreach (XmlNode node1 in node.ChildNodes)
                    {
                        XmlElement ele1 = (XmlElement)node1;
                        if (ele1.GetAttribute("ID") == "64655")
                        {
                            ele1.SetAttribute("ID", item.Tube_ID);

                        }
                    }
                }
                //Order ID
                XmlNode xn = xmlDoc.SelectSingleNode("Order");

                XmlElement ele2 = (XmlElement)xn;

                ele2.SetAttribute("OrderId", item.OrderName);
                string filenamesave = file + "\\" + item.OrderName + "-" + item.Tube_ID + ".order.xml";

                xmlDoc.Save(filenamesave);//设置一个保存路径 
            }
        }

        private void NewMethod()
        {
            if (this.dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("当前界面没有数据，请确认 !", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "csv|*.csv";
            string strFileName = "Data " + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            FileStream fa = new FileStream(strFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa, Encoding.Unicode);
            string delimiter = "\t";
            string strHeader = "";
            for (int i = 0; i < this.dataGridView1.Columns.Count; i++)
            {
                strHeader += this.dataGridView1.Columns[i].HeaderText + delimiter;
            }
            sw.WriteLine(strHeader);

            //output rows data
            for (int j = 0; j < this.dataGridView1.Rows.Count; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < this.dataGridView1.Columns.Count; k++)
                {
                    if (this.dataGridView1.Rows[j].Cells[k].Value != null)
                    {
                        strRowValue += this.dataGridView1.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ").Replace("\n", "") + delimiter;
                        if (this.dataGridView1.Rows[j].Cells[k].Value.ToString() == "LIP201507-35")
                        {

                        }

                    }
                    else
                    {
                        strRowValue += this.dataGridView1.Rows[j].Cells[k].Value + delimiter;
                    }
                }
                sw.WriteLine(strRowValue);
            }
            sw.Close();
            fa.Close();
            MessageBox.Show("下载完成！", "保存", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var shops = this.Results.Where(s => s.Rack_ID != null && s.Rack_ID.ToString().StartsWith(textBox1.Text.ToString())).ToList();
            if (shops.Count != 0)
            {
                var store = shops.First();
                this.comboBox1.SelectedValue = store.Rack_ID;

            }
        }

        private void notifyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<xmlDataSources> orders = new List<xmlDataSources>();
            for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
            {
                var order = dataGridView1.SelectedRows[i].DataBoundItem as xmlDataSources;
                orders.Add(order);
            }
            clsAllnew BusinessHelp = new clsAllnew();

            BusinessHelp.delete_XMLServer(orders);
            InitializeDataSource();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确认要清空所有数据（请谨慎操作） ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {

            }
            else
                return;
            clsAllnew BusinessHelp = new clsAllnew();

            BusinessHelp.deleteall_XMLServer();
            InitializeDataSource();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            InitializeDataSource();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            var form = new Importxml();
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            {
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewColumn column = dataGridView1.Columns[e.ColumnIndex];

            if (column == deleteColumn1)
            {
                List<xmlDataSources> orders = new List<xmlDataSources>();

                var row = dataGridView1.Rows[e.RowIndex];
                var model = row.DataBoundItem as xmlDataSources;
                string msg = string.Format("确定删除<{0}>？", model.Tube_ID);
                if (MessageBox.Show(msg, "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    orders.Add(model);
                    clsAllnew BusinessHelp = new clsAllnew();
                    BusinessHelp.delete_XMLServer(orders);
                    InitializeDataSource();
                }
                else
                    return;

            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
                return;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (checkedav == false)
                    dataGridView1.Rows[i].Cells[0].Value = true;
                else
                    dataGridView1.Rows[i].Cells[0].Value = false;

            }
            if (checkedav == false)
            {
                dataGridView1.Rows[0].Cells[0].Value = true;
                checkedav = true;

            }
            else
            {
                dataGridView1.Rows[0].Cells[0].Value = false;
                checkedav = false;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            clsAllnew BusinessHelp = new clsAllnew();

            List<xmlDataSources> allServer = BusinessHelp.findALLXML_Server();

            var list = allServer.Select(p => p.Rack_ID).Distinct().ToList();

            BusinessHelp.status_Server(list);

        }

    }
}
