using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace KM_BiotechnologyXML
{
    public partial class MainForm : Form
    {
        private Thread GetDataforRawDataThread;
        public log4net.ILog ProcessLogger;
        public log4net.ILog ExceptionLogger;
        Sunisoft.IrisSkin.SkinEngine se = null;
        private frmDataCenter frmDataCenter;
        private System.Timers.Timer timerAlter1;
        private bool IsRun1 = false;
        frmAboutBox aboutbox;
        private int alterisrun;
        public MainForm()
        {
            InitializeComponent();
            InitialSystemInfo();
            aboutbox = new frmAboutBox();
            #region Noway
            DateTime oldDate = DateTime.Now;
            DateTime dt3;
            string endday = DateTime.Now.ToString("yyyy/MM/dd");
            dt3 = Convert.ToDateTime(endday);
            DateTime dt2;
            dt2 = Convert.ToDateTime("2017/06/18");

            TimeSpan ts = dt2 - dt3;
            int timeTotal = ts.Days;
            if (timeTotal < 0)
            {
                MessageBox.Show("Please Contact your administrator !");
                toolStripDropDownButton2.Enabled = false;
                toolStripDropDownButton1.Enabled = false;

                this.Close();

                return;
            }
            #endregion

        }
        private void InitialSystemInfo()
        {
            #region 初始化配置
            ProcessLogger = log4net.LogManager.GetLogger("ProcessLogger");
            ExceptionLogger = log4net.LogManager.GetLogger("SystemExceptionLogger");

            #endregion
        }
        private void pBBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var form = new Importxml();
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            {
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (frmDataCenter == null)
            {
                frmDataCenter = new frmDataCenter();
                frmDataCenter.FormClosed += new FormClosedEventHandler(FrmOMS_FormClosed);
            }
            if (frmDataCenter == null)
            {
                frmDataCenter = new frmDataCenter();
            }
            frmDataCenter.Show(this.dockPanel2);


        }
        void FrmOMS_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (sender is frmDataCenter)
            {
                frmDataCenter = null;
            }
        }

        private void 一键配置ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 关于系统ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            aboutbox.ShowDialog();
        }

        private void 一键配置初始化信息ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

                //安装数据库
                systemin();
                ProcessLogger.Fatal("10103--Create amdin Start" + DateTime.Now.ToString());


                NewMethod1();

                MessageBox.Show("导入成功,可以使用了！");

            }
            catch (Exception ex)
            {
                MessageBox.Show("导入数据错误，请确认本地文件包解压是否正常" + ex, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

                throw;
            }

        }
        private void systemin()
        {

            try
            {
                #region 创建文件夹和 log 记事本

                ProcessLogger.Fatal("1001--Create Folder txt" + DateTime.Now.ToString());
                string spath = @"C:\Program Files\mongodb\bin";

                if (Directory.Exists(spath))
                {
                }
                else
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(spath);
                    directoryInfo.Create();
                }


                spath = @"C:\Program Files\mongodb\data\db";

                if (Directory.Exists(spath))
                {
                }
                else
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(spath);
                    directoryInfo.Create();
                }

                spath = @"C:\Program Files\mongodb\data\log";

                if (Directory.Exists(spath))
                {
                }
                else
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(spath);
                    directoryInfo.Create();
                }
                spath = @"C:\Program Files\mongodb\data\log\MongoDB.log";

                if (File.Exists(spath))
                {
                }
                else
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(spath);

                    System.IO.File.Create(spath);
                }

                #endregion
                #region 复制文件BIN 到指定目录
                ProcessLogger.Fatal("1002--copy bin" + DateTime.Now.ToString());
                string srcdir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\bin");
                string todir = @"C:\Program Files\mongodb\";
                string dstdir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\bin");
                bool overwrite = true;
                CopyDirIntoDestDirectory(srcdir, todir, overwrite);


                #endregion

                #region 调用CMD 命令
                ProcessLogger.Fatal("1003--install db Start" + DateTime.Now.ToString());
                string cmd = @"C:&cd C:\Program Files\mongodb\bin&&mongod --dbpath ""C:\Program Files\mongodb\data\db""";
                string output = "";
                //cmd = @"ipconfig/all";
                RunCmd(cmd, out output);
                //  MessageBox.Show(output);

                ProcessLogger.Fatal("1004--install servers" + DateTime.Now.ToString());
                timerAlter1 = new System.Timers.Timer(200000);
                timerAlter1.Elapsed += new System.Timers.ElapsedEventHandler(TimeControl1);
                timerAlter1.AutoReset = true;
                timerAlter1.Start();
                cmd = @"C:&cd C:\Program Files\mongodb\bin&&mongod --dbpath ""C:\Program Files\mongodb\data\db"" --logpath ""C:\Program Files\mongodb\data\log\MongoDB.log"" --install --serviceName ""MongoDB""";
                RunCmd(cmd, out output);
                #endregion

                //MessageBox.Show("运行结束 ，后台数据配置成功 ", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                if (ex.Message.ToString().Contains("AccessException"))
                {
                    string dstdir = "";
                    Version ver = System.Environment.OSVersion.Version;
                    if (ver.Major.ToString().Contains("10"))
                    {
                        dstdir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\win10Admin.reg");
                    }
                    else if (ver.Major.ToString().Contains("6"))
                    {
                        dstdir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\win7Admin.reg");
                    }
                    Process.Start(dstdir);

                }
                MessageBox.Show("10901:由于您未获得管理员权限，请尝试取得管理员权限\r\n（系统(仅支持Window10，win7版本)已自动尝试获取权限，如重试启动系统还未正常运行则请手动获取windows 的权限） ！" + ex, "安装错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

                throw;
            }

        }
        public static void CopyDirIntoDestDirectory(string srcdir, string dstdir, bool overwrite)
        {
            string todir = Path.Combine(dstdir,
                                        Path.GetFileName(srcdir)
                                        );

            if (!Directory.Exists(todir))
                Directory.CreateDirectory(todir);

            foreach (var s in Directory.GetFiles(srcdir))
            {
                string news = s.Replace(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\bin"), todir);
                if (File.Exists(news))
                {
                }
                else
                {
                    File.Copy(s, Path.Combine(todir, Path.GetFileName(s)), overwrite);
                }
            }
            foreach (var s in Directory.GetDirectories(srcdir))
                CopyDirIntoDestDirectory(s, todir, overwrite);
        }
        private void TimeControl1(object sender, EventArgs e)
        {
            if (!IsRun1)
            {
                IsRun1 = true;
                GetDataforRawDataThread = new Thread(TimeMethod1);
                GetDataforRawDataThread.Start();
            }
        }
        public static void RunCmd(string cmd, out string output)
        {
            try
            {
                string CmdPath = @"C:\Windows\System32\cmd.exe";
                cmd = cmd.Trim().TrimEnd('&') + "&exit";//说明：不管命令是否成功均执行exit命令，否则当调用ReadToEnd()方法时，会处于假死状态
                using (Process p = new Process())
                {
                    p.StartInfo.FileName = CmdPath;
                    p.StartInfo.UseShellExecute = false;        //是否使用操作系统shell启动
                    p.StartInfo.RedirectStandardInput = true;   //接受来自调用程序的输入信息
                    p.StartInfo.RedirectStandardOutput = true;  //由调用程序获取输出信息
                    p.StartInfo.RedirectStandardError = true;   //重定向标准错误输出
                    p.StartInfo.CreateNoWindow = true;          //不显示程序窗口
                    p.Start();//启动程序

                    //向cmd窗口写入命令
                    p.StandardInput.WriteLine(cmd);
                    p.StandardInput.AutoFlush = true;

                    //获取cmd窗口的输出信息
                    output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();//等待程序执行完退出进程
                    p.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("EX:数据库配置失败 ：" + ex);


                throw;
            }
        }

        private void TimeMethod1()
        {
            string output = "";
            string cmd = @"C:&cd C:\Program Files\mongodb\bin&&mongod --dbpath ""C:\Program Files\mongodb\data\db"" --logpath ""C:\Program Files\mongodb\data\log\MongoDB.log"" --install --serviceName ""MongoDB""";
            RunCmd(cmd, out output);

            alterisrun = 0;
            IsRun1 = false;
            MessageBox.Show("运行结束 ，后台数据配置成功 ,系统即将关闭，请自行重启即可", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();

        }

        private void NewMethod1()
        {


        }

        private void 打开本地目录ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System"), "");
            System.Diagnostics.Process.Start("explorer.exe", ZFCEPath);
        }

        private void 追踪分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ""), "");
            System.Diagnostics.Process.Start("系统使用说明.xls", ZFCEPath);
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
        }

        private void MainForm_Deactivate(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Visible = false;
            }
        }


    }
}
