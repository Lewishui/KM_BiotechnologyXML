using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using clsdatabaseinfo;
using clsKMBuiness;

namespace KM_BiotechnologyXML
{
    public partial class Importxml : Form
    {
        string filepath;
        List<xmlDataSources> Results;
        string filename;

        public Importxml()
        {
            InitializeComponent();
        }

        private void openFileBtton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                pathTextBox.Text = dialog.SelectedPath;
                //MessageBox.Show("已选择文件夹:" + foldPath, "选择文件夹提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pathTextBox.Text = @"C:\ProgramData\TTPLabTach\arktic\Roports";
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            filepath = pathTextBox.Text;

            if (pathTextBox.Text == null || pathTextBox.Text == "")
            {
                MessageBox.Show("请选择路径!");

                return;
            }
            this.importButton.Enabled = false;
            this.cancelButton.Enabled = true;
            this.closeButton.Enabled = false;
            if (backgroundWorker1.IsBusy != true)
            {
                backgroundWorker1.RunWorkerAsync(new WorkerArgument { OrderCount = 0, CurrentIndex = 0 });

            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            //bool success = Execute(pathTextBox.Text, worker, e);
            bool success = NewMethod(worker, e);
        }
        private bool NewMethod(BackgroundWorker worker, DoWorkEventArgs e)
        {
            WorkerArgument arg = e.Argument as WorkerArgument;
            int progress = 0;
            bool success = true;
            try
            {
                Results = new List<xmlDataSources>();

                List<string> Alist = GetBy_CategoryReportFileName(filepath);
                arg.OrderCount = Alist.Count;
                for (int i = 0; i < Alist.Count; i++)
                {
                    filename = Alist[i];

                    progress = Convert.ToInt16(((i) * 1.0 / Alist.Count) * 100);

                    arg.CurrentIndex = i;

                    LoadSalesData(filepath + "\\" + Alist[i]);

                    backgroundWorker1.ReportProgress(progress, arg);
                }
                //写入数据库
                clsAllnew BusinessHelp = new clsAllnew();

                BusinessHelp.SPInputclaimreport_Server(Results);
                backgroundWorker1.ReportProgress(100, arg);
                e.Result = string.Format("{0} 条正常导入成功", Results.Count);

            }
            catch (Exception ex)
            {
                if (!e.Cancel)
                {
                    //arg.HasError = true;
                    //arg.ErrorMessage = exception.Message;
                    e.Result = ex.Message + "或所导入信息超出要求长度";
                }
                success = false;
            }

            return success;
        }
        private void LoadSalesData(string path)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(path);
                XmlElement xmlRoot = xmlDoc.DocumentElement;
              
                //new
                ReadNewMethod(xmlRoot, path, xmlDoc);
                //end 
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private void ReadNewMethod(XmlElement xmlRoot, string path, XmlDocument xmlDoc)
        {
            XmlNode xn = xmlDoc.SelectSingleNode("RackReport");
            string OrderName = ((XmlElement)xn).GetAttribute("OrderName");   //获取Name属性值 

            {
                //string id00 = ((XmlElement)node00).GetAttribute("OrderName");   //获取Name属性值 


                XmlNodeList personNodes0 = xmlRoot.GetElementsByTagName("Rack"); //获取Person子节点集合 
                foreach (XmlNode node0 in personNodes0)
                {

                    string id0 = ((XmlElement)node0).GetAttribute("ID");   //获取Name属性值 

                    XmlNodeList personNodes = xmlRoot.GetElementsByTagName("Hole"); //获取Person子节点集合 
                    foreach (XmlNode node in node0.ChildNodes)
                    {
                        string id = ((XmlElement)node).GetAttribute("ID");   //获取Name属性值 
                        //  string name = ((XmlElement)node).GetElementsByTagName("ID")[0].InnerText;
                        foreach (XmlNode node1 in node.ChildNodes)
                        {
                            string id1 = ((XmlElement)node1).GetAttribute("ID");
                            foreach (XmlNode node2 in node1.ChildNodes)
                            {
                                string id2 = ((XmlElement)node2).GetAttribute("ID");
                                xmlDataSources tempnote = new xmlDataSources(); //定义返回值
                                tempnote.Rack_ID = id0;
                                tempnote.Hole_ID = id1;
                                tempnote.Tube_ID = id2;
                                tempnote.Input_Date = DateTime.Now.ToString("yyyyMMdd");
                                tempnote.OrderName = OrderName;
                                tempnote.FileName = filename;
                                Results.Add(tempnote);


                            }
                        }

                    }
                }
            }
        }

        private List<string> GetBy_CategoryReportFileName(string dirPath)
        {

            List<string> FileNameList = new List<string>();
            ArrayList list = new ArrayList();

            if (Directory.Exists(dirPath))
            {
                list.AddRange(Directory.GetFiles(dirPath));
            }
            if (list.Count > 0)
            {
                foreach (object item in list)
                {
                    if (!item.ToString().Contains("~$") && item.ToString().Contains("xml"))
                        FileNameList.Add(item.ToString().Replace(dirPath + "\\", ""));
                }
            }

            return FileNameList;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            WorkerArgument arg = e.UserState as WorkerArgument;
            if (!arg.HasError)
            {
                this.progressMsgLabel.Text = String.Format("{0}/{1}", arg.CurrentIndex, arg.OrderCount);
                this.ProgressValue = e.ProgressPercentage;
            }
            else
            {
                this.progressMsgLabel.Text = arg.ErrorMessage;
            }
        }
        public int ProgressValue
        {
            get { return this.progressBar1.Value; }
            set { progressBar1.Value = value; }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.cancelButton.Enabled = false;
            this.closeButton.Enabled = true;
            this.importButton.Enabled = true;

            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show(string.Format("It is cancelled!"));
            }
            else
            {
                MessageBox.Show(string.Format("{0}", e.Result));

            }

        }
    }
}
