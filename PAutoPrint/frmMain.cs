using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.IO.Ports;

namespace PAutoPrintReport
{
    public partial class frmMain : Form
    {
        #region Constant, Struct and Enum
        enum _StepProcess : int
        {
            InitialDatabase = 0
                ,
            InitialPrint = 1
                ,
            InitialComportPLC = 2
                ,
            InitialDataGrid = 3
                ,
            InitialClassEvent = 4
                ,
            ChangeProcess = 5
        }

        #endregion
        
        public Logfile LogFile = new Logfile();

        private PrintProcess[] PrintProcess;
        private PrintProcessWeight[] PrintProcessWeight;
        private PrintInstructionProcess[] PrintInstructionProcess;
        private PrintEndOfDayProcess[] PrintEndOfDayProcess;
        public Database OraDb;
        public IniLib.CINI iniFile = new IniLib.CINI();
        string arg;
        string[] args;
        string scanID;
        string logFileName;
        int processID;
        public string App_Path;

        _StepProcess mainStepProcess;

        private bool IsSingleInstance()
        {
            foreach (Process process in Process.GetProcesses())
            {
                if (process.MainWindowTitle == this.Text)
                    return false;
            }
            return true;
        }

        public frmMain()
        {
            InitializeComponent();
            APP_PATH();
            try
            {
                args = Environment.GetCommandLineArgs();
                arg = args[1];
            }
            catch (Exception exp)
            {
                arg = "1";
                //AddListBoxItem(mBayNo);
            }
            this.Text = iniFile.INIRead(Directory.GetCurrentDirectory() + "\\AppStartup.ini", arg, "TITLE");
            scanID = iniFile.INIRead(Directory.GetCurrentDirectory() + "\\AppStartup.ini", arg, "SCANID");
            logFileName = this.Text;
            if (!IsSingleInstance())
            {
                MessageBox.Show("Another instance of this app is running.", this.Text);
                //Application.Exit();
                System.Environment.Exit(1);
            }
            //StartThread();
            AddListBox = DateTime.Now + "><------------Application Start------------->";

        }

        #region Thread
        bool thrConnect;
        bool thrShutdown;
        bool thrRunning;

        Thread thrMain;

        private void StartThread()
        {
            //System.Threading.Thread.Sleep(1000);
            thrMain = new Thread(this.RunProcess);
            thrMain.Name = this.Text;
            thrMain.Start();
        }

        private void RunProcess()
        {
            thrRunning = true;
            System.Threading.Thread.Sleep(1000);
            if (mainStepProcess != _StepProcess.ChangeProcess)
            {
                InitialDataBase();
                mainStepProcess = _StepProcess.InitialDatabase;
            }
            AddListBox = "Initial report path " + App_Path.ToString() + "\\Report File\\";
            while (thrRunning)
            {
                try
                {
                    if (thrShutdown)
                        return;
                    switch (mainStepProcess)
                    {
                        case _StepProcess.InitialDatabase:
                            if (OraDb.ConnectStatus())
                            {
                                mainStepProcess = _StepProcess.InitialPrint;
                                //AddListBoxItem(mStepProcess.ToString());
                            }

                            Thread.Sleep(500);
                            break;
                        case _StepProcess.InitialPrint:
                            Thread.Sleep(500);
                            InitialDataGrid();
                            if (InitialPrintProcess())
                            {
                                mainStepProcess = _StepProcess.InitialDataGrid;
                                PrintProcess[processID].StartThread();
                            }

                            if (InitialPrintInstructionProcess())
                            {
                                mainStepProcess = _StepProcess.InitialDataGrid;
                                PrintInstructionProcess[processID].StartThread();
                            }
                            if (InitialPrintProcessWeight())
                            {
                                mainStepProcess = _StepProcess.InitialDataGrid;
                                PrintProcessWeight[processID].StartThread();
                            }
                            if (InitialPrintEndOfDayProcess())
                            {
                                mainStepProcess = _StepProcess.InitialDataGrid;
                                PrintEndOfDayProcess[processID].StartThread();
                            }
                            break;
                        case _StepProcess.InitialDataGrid:
                            //thrRunning = false;
                           // InitialDataGrid();
                            //AddListBoxItem(mStepProcess.ToString());
                            mainStepProcess = _StepProcess.InitialClassEvent;
                            break;
                        case _StepProcess.InitialClassEvent:
                            //thrRunning = false;
                            thrShutdown = true;
                            InitialClassEvent();
                            break;
                        case _StepProcess.ChangeProcess:
                            thrShutdown = true;
                            Thread.Sleep(300);
                            while (PrintProcess[processID].IsThreadAlive)
                            { }
                            //PrintProcess[processID].StartThread();
                            break;
                    }
                    DisplayDateTime();
                    Thread.Sleep(300);
                }
                catch (Exception exp)
                { AddListBoxItem(DateTime.Now + ">" + exp.Message + "[" + exp.Source + "-" + mainStepProcess.ToString() + "]"); }
                //finally
                //{
                //    mShutdown = true;
                //    mRunning = false;
                //}
            }
        }
        #endregion

        #region ListboxItem

        public void DisplayMessage(string pFileName, string pMsg)
        {
            if (this.lstMain.InvokeRequired)
            {
                // This is a worker thread so delegate the task.
                if (lstMain.Items.Count > 1000)
                {
                    //lstMain.Items.Clear();
                    this.Invoke((Action)(() => lstMain.Items.Clear()));
                }

                this.lstMain.Invoke(new DisplayMessageDelegate(this.DisplayMessage), pFileName, pMsg);
            }
            else
            {
                // This is the UI thread so perform the task.
                if (pMsg != null)
                {
                    if (lstMain.Items.Count > 1000)
                    {
                        //lstMain.Items.Clear();
                        this.Invoke((Action)(() => lstMain.Items.Clear()));
                    }

                    this.lstMain.Items.Insert(0, pMsg);
                    //logfile.WriteLog("System", item.ToString());
                    //PLog.WriteLog(pFileName, iMsg);
                }
            }
        }

        private delegate void DisplayMessageDelegate(string pFileName, string iMsg);

        public object AddListBox
        {
            set
            {
                AddListBoxItem(value);
            }
        }

        private delegate void AddListBoxItemDelegate(object item);

        private void AddListBoxItem(object pItem)
        {
            if (this.lstMain.InvokeRequired)
            {
                // This is a worker thread so delegate the task.
                if (lstMain.Items.Count > 1000)
                {
                    //lstMain.Items.Clear();
                    this.Invoke((Action)(() => lstMain.Items.Clear()));
                }

                this.lstMain.Invoke(new AddListBoxItemDelegate(this.AddListBoxItem), pItem);
            }
            else
            {
                // This is the UI thread so perform the task.
                if (pItem != null)
                {
                    if (lstMain.Items.Count > 1000)
                    {
                        //lstMain.Items.Clear();
                        this.Invoke((Action)(() => lstMain.Items.Clear()));
                    }

                    this.lstMain.Items.Insert(0, (processID + 1).ToString() + "-" + logFileName + ">" + pItem);
                    //logfile.WriteLog("System", item.ToString());
                    LogFile.WriteLog(logFileName, (processID + 1).ToString() + "-" + logFileName + ">" + pItem.ToString());
                }
            }
        }

        private delegate void ClearListBoxItemDelegate();
        private void ClearListBoxItem()
        {
            if (this.lstMain.InvokeRequired)
            {
                this.lstMain.Invoke(new ClearListBoxItemDelegate(ClearListBoxItem));

            }
            else
            {
                this.lstMain.Items.Clear();
            }

        }
        #endregion


        #region Combobox
        private delegate void AddComboboxItemEvenHandler(object pItem);

       
        #endregion

        #region MynotifyIcon
        private void FormResize()
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                mynotifyIcon1.Icon = this.Icon;
                mynotifyIcon1.Visible = true;
                mynotifyIcon1.BalloonTipText = this.Text;
                mynotifyIcon1.ShowBalloonTip(500);
                this.Hide();
            }
        }
        private void mynotifyIcon1_Click(object sender, MouseEventArgs e) 
        {
            mynotifyIcon1.Visible = false;
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }
        #endregion

        #region Class Events
        void InitialATGEventHandler()
        {
            //ATGProcess.ATGEventsHaneler handler1 = new ATGProcess.ATGEventsHaneler(WriteEventsHandler);
            //atgProcess[0].OnATGEvents += handler1;
        }

        void InitialComportEventHandler()
        {
            Comport.ComportEventsHandler hander1 = new Comport.ComportEventsHandler(WriteEventsHandler);
            //for (int i = 0; i < ATGComport.Length; i++)
            //{
            //    ATGComport[i].OnComportEvents += hander1;
            //}
        }

        void WriteEventsHandler(object pSender, string pMessage)
        {
            AddListBoxItem(pMessage);
            //mLog.WriteLog(mLogFileName, message);
        }
        #endregion

        #region Main Step Process
        private void InitialDataBase()
        {
            OraDb = new Database(this);
        }

        private bool InitialPrintProcess()
        {
           
            bool vRet = false;
            PrintProcess = new PrintProcess[1];
            PrintProcess[0] = new PrintProcess(this,1);
            vRet = true;
            return vRet;
        }

        private bool InitialPrintInstructionProcess()
        {

            bool vRet = false;
            PrintInstructionProcess = new PrintInstructionProcess[1];
            PrintInstructionProcess[0] = new PrintInstructionProcess(this, 1);
            vRet = true;
            return vRet;
        }

        private bool InitialPrintProcessWeight()
        {

            bool vRet = false;
            PrintProcessWeight = new PrintProcessWeight[1];
            PrintProcessWeight[0] = new PrintProcessWeight(this, 1);
            vRet = true;
            return vRet;
        }
        private bool InitialPrintEndOfDayProcess()
        {

            bool vRet = false;
            PrintEndOfDayProcess = new PrintEndOfDayProcess[1];
            PrintEndOfDayProcess[0] = new PrintEndOfDayProcess(this, 1);
            vRet = true;
            return vRet;
        }
       

        private void InitialClassEvent()
        {
            //InitialATGEventHandler();
            //InitialComportEventHandler();
        }

        public void ChangeProcess()
        {
            processID = 1;
            //plcProcess[processID].StartThread();
            thrShutdown = false;
            mainStepProcess = _StepProcess.ChangeProcess;
            StartThread();
        }
        #endregion

        private void DisplayDateTime()
        {
            toolStripStatusLabel1.Text = "Database connect = " + OraDb.ConnectStatus().ToString() +
                                                "[" + OraDb.ConnectServiceName() + "]" +
                                                "   [Date Time : " + DateTime.Now + "]";
        }

        

        private void frmMain_Resize(object sender, EventArgs e)
        {
            FormResize();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                toolStripStatusLabel1.Text = "Database connect = " + OraDb.ConnectStatus().ToString() +
                                                "[" + OraDb.ConnectServiceName() + "]" +
                                                "   [Date Time : " + DateTime.Now + "]";

            }
            catch (Exception exp)
            { }   
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            AddListBox = DateTime.Now + "><------------Application Stop-------------->";
            thrShutdown = true;
            iniFile = null;
            //VisibleObject(true);
            timer1.Enabled = false;
            Application.DoEvents();
            this.Cursor = Cursors.WaitCursor;
            //System.Threading.Thread.Sleep(500);
            //if(ATGComport!=null)
            //{
            //    foreach (Comport p in ATGComport)
            //    {
            //        p.Dispose();
            //    }
            //}
            
            System.Threading.Thread.Sleep(300);
            if (PrintProcess != null)
            {
                for (int i = 0; i < PrintProcess.Length; i++)
                {
                    PrintProcess[i].StopThread();
                    PrintProcess[i].Dispose();
                }
            }

            if(PrintProcessWeight != null)
            {
                for (int i = 0; i < PrintProcessWeight.Length; i++)
                {
                    PrintProcessWeight[i].StopThread();
                    PrintProcessWeight[i].Dispose();
                }
            }

            if (PrintInstructionProcess != null)
            {
                for (int i = 0; i < PrintInstructionProcess.Length; i++)
                {
                    PrintInstructionProcess[i].StopThread();
                    PrintInstructionProcess[i].Dispose();
                }
            }

            if (PrintEndOfDayProcess != null)
            {
                for (int i = 0; i < PrintEndOfDayProcess.Length; i++)
                {
                    PrintEndOfDayProcess[i].StopThread();
                    PrintEndOfDayProcess[i].Dispose();
                }
            }
            OraDb.Close();
            OraDb.Dispose();         
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            StartThread();
            timer1.Enabled = true;
        }

        private void lstMain_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            MessageBox.Show(lstMain.SelectedItem.ToString(),this.Text,MessageBoxButtons.OK);
        }

        #region DataGrid
        private delegate void AddDataGridItemEventHandler(int pRow, int pCol, object pValue);
        private delegate void AddDataGridRowsEventHandler(int pRows);

        private void AddDataGridRows(int pRows)
        {
            if (this.dataGridView1.InvokeRequired)
            {
                // This is a worker thread so delegate the task.

                this.dataGridView1.Invoke(new AddDataGridRowsEventHandler(this.AddDataGridRows), pRows);
            }
            else
            {
                // This is the UI thread so perform the task.
                if (pRows != 0)
                {
                    dataGridView1.Rows.Add(pRows);
                }
            }
        }
        private void AddDataGridItem(int pRow, int pCol, object pValue)
        {
            if (this.dataGridView1.InvokeRequired)
            {
                // This is a worker thread so delegate the task.

                this.dataGridView1.Invoke(new AddDataGridItemEventHandler(this.AddDataGridItem), pRow, pCol, pValue);
            }
            else
            {
                // This is the UI thread so perform the task.
                if (pRow >= 0)
                {
                    //dataGridView1.Rows[1].Cells[0].Value = "bay_no";
                    dataGridView1.Rows[pRow].Cells[pCol].Value = pValue;
                }
            }
        }
        #endregion
        private void InitialDataGrid()
        {
                    //dataGridView1.Rows.Add(dt.Rows.Count - 1);
                    AddDataGridRows(2);

                        AddDataGridItem(0, 0, 1);
                        AddDataGridItem(0, 1, "AutoPrint");
                        AddDataGridItem(0, 2, "Loading Introduction(TM)");
                        AddDataGridItem(1, 0, 2);
                        AddDataGridItem(1, 1, "AutoPrint");
                        AddDataGridItem(1, 2, "Delivery Receipt");
                        //AddDataGridItem(2, 0, 3);
                        //AddDataGridItem(2, 1, "AutoPrint");
                        //AddDataGridItem(2, 2, "ใบชั่งน้ำหนัก");
                        AddDataGridItem(2, 0, 3);
                        AddDataGridItem(2, 1, "AutoPrint");
                        AddDataGridItem(2, 2, "End Of Day");
                 
        }
        public void APP_PATH()
        {
            App_Path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        }
    }
       
}
