using System;
using System.Management;
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
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Drawing.Printing;

namespace PAutoPrintReport
{
    public class PrintInstructionProcess : IDisposable
    {
         #region Constant, Struct and Enum
        const int offsetFC01 = 1;
        const int offsetFC02 = 100001;
        const int offsetFC03 = 400001;
        const int offsetFC04 = 300001;
        const int offsetFC05 = 1;
        const int offsetFC06 = 400001;
        const ushort offsetFC15 = 1;
        const int offsetFC16 = 400001;
        private string LoadHeader_NO;
        private string PrintStatus;
        private string out_ReportID;
        private string out_ReportPath;
        private string out_PrinterID;
        private string out_PrinterName;
        private string out_Parametername;
               
        struct _PrintBuffer
        {
            public int Id;
            public string ReportID;
            public string ReportName;
            public string ReportPath;
            public string PrinterName;
            public string ParamterName;
            public bool AutoPrint;
            public DateTime TimeStamp;
            public int PrintStatus;
            public string LoadHeader_NO;
        }
                
        enum _PrintInstructionStepProcess :int
        {
            InitialReport=0,
            CheckStatusReport=1,
            LoadReport=10,
            UpdateStatus=11,
            ChangeProcess=30
        } 
        #endregion
        int i = 1;
       frmMain fMain;
       _PrintInstructionStepProcess PrintInstructionStepProcess;

        int processId=1;
        //int atgAddress;
        DateTime chkResponse;

        #region Construct and Deconstruct
        private bool IsDisposed = false;
        
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
            //mRunn = false;
        }
        
        protected void Dispose(bool Diposing)
        {
            if (!IsDisposed)
            {
                if (Diposing)
                {
                    //Clean Up managed resources
                    thrShutdown = true; 
                }
                //Clean up unmanaged resources
            }
            IsDisposed = true;
        }
        
        public PrintInstructionProcess(frmMain pFrm)
        {
            fMain = pFrm;
        }

        public PrintInstructionProcess(frmMain pFrm, int pId)
        {
            fMain = pFrm;
            processId = pId;  
        }

        ~PrintInstructionProcess()
        { }
        #endregion

        #region Class Events
        public delegate void ATGEventsHaneler(object pSender, string pEventMsg);
        string logFileName;

        void RaiseEvents(string pSender, string pMsg)
        {
            string vMsg = DateTime.Now + ">[" + pSender + "]" + pMsg;
            logFileName = "Auto Print Instroduction LLTLB";
            try
            {
                fMain.AddListBox = "" + logFileName + ">" + vMsg;
                //fMain.LogFile.WriteLog(logFileName, vMsg);
            }
            catch (Exception exp)
            { }
        }

        void RaiseEvents(string pMsg)
        {
            logFileName = "Auto Print Instroduction LLTLB";
            try
            {
                fMain.AddListBox = DateTime.Now + "> " + pMsg;
                //fMain.LogFile.WriteLog(logFileName, vMsg);
            }
            catch (Exception exp)
            { }
        }
        #endregion

        #region Thread
        bool thrConnect;
        bool thrShutdown;
        bool thrRunning;

        Thread thrMain;

        public void StartThread()
        {
            System.Threading.Thread.Sleep(1000);
            thrMain = new Thread(this.RunProcess);
            thrMain.Name = processId.ToString() + "thrPrintInstruction";
            thrMain.Start();
        }

        public void StopThread()
        {
            thrShutdown = true;
        }

        private void RunProcess()
        {
            thrRunning = true;
            thrShutdown = false;
            PrintInstructionStepProcess = _PrintInstructionStepProcess.InitialReport;
            chkResponse = DateTime.Now;
            while (thrRunning)
            {
                try  
                {
                    if (thrShutdown)
                        return;
                    switch (PrintInstructionStepProcess)
                    {
                        case _PrintInstructionStepProcess.InitialReport:
                            if (GetConfigReport())
                            {
                                PrintInstructionStepProcess = _PrintInstructionStepProcess.CheckStatusReport;     
                            }
                            break;
                        case _PrintInstructionStepProcess.CheckStatusReport:
                            if (CheckStatusInstructionReport())
                            {
                                PrintInstructionStepProcess = _PrintInstructionStepProcess.LoadReport;     
                            }
                            break;
                        case _PrintInstructionStepProcess.LoadReport:
                            if (LoadInstructionReport())
                            {
                                UpdateInstructionToDatabase(LoadHeader_NO);
                                PrintInstructionStepProcess = _PrintInstructionStepProcess.ChangeProcess;
                            }
                            else
                            {
                                PrintInstructionStepProcess = _PrintInstructionStepProcess.CheckStatusReport;
                            }
                            break;
                        case _PrintInstructionStepProcess.UpdateStatus:
                           // UpdateInstructionToDatabase(LoadHeader_NO);
                            PrintInstructionStepProcess = _PrintInstructionStepProcess.ChangeProcess; ;
                            break;
                        case _PrintInstructionStepProcess.ChangeProcess:
                            PrintInstructionStepProcess = _PrintInstructionStepProcess.CheckStatusReport;
                            break;
                    }
                    Thread.Sleep(1000);
                }
                catch (Exception exp)
                { 
                    RaiseEvents(exp.Message + "[source>" + exp.Source + "]");
                    Thread.Sleep(3000);
                }
                //finally
                //{
                //    mShutdown = true;
                //    mRunning = false;
                //}
                Thread.Sleep(500);
            }
        }
        #endregion

        #region define Enum String Value
        enum _DataType
        {
            [EnumValue("BOOLEAN")]
            BOOLEAN = 1,
            [EnumValue("BYTE")]
            BYTE = 2,
            [EnumValue("SHORT INT")]
            SHORT_INT = 3,
            [EnumValue("LONG INT")]
            LONG_INT = 4,
            [EnumValue("FLOAT")]
            FLOAT = 5,
            [EnumValue("STRING")]
            STRING = 6,
        }
        class EnumValue : System.Attribute
        {
            private string _value;
            public EnumValue(string value)
            {
                _value = value;
            }
            public string Value
            {
                get { return _value; }
            }
        }
        static class EnumString
        {
            public static string GetStringValue(Enum value)
            {
                string output = null;
                Type type = value.GetType();
                System.Reflection.FieldInfo fi = type.GetField(value.ToString());
                EnumValue[] attrs = fi.GetCustomAttributes(typeof(EnumValue), false) as EnumValue[];
                if (attrs.Length > 0)
                {
                    output = attrs[0].Value;
                }
                return output;
            }
        } 
        #endregion

        #region Change process or comport
        void ChangeProcess()
        {
            Thread.Sleep(3000);
            bool vRet = false;
           
            var vDiff = (DateTime.Now - chkResponse).TotalMinutes;
            if ((vRet) && (vDiff >= 1))
            {
                chkResponse = DateTime.Now;
                thrShutdown = true;
                //ChangeComport();
                RaiseEvents("Communication changed.");
                fMain.ChangeProcess();
            }   
        }

        public bool IsThreadAlive
        {
            get { return thrMain.IsAlive; }
        }

        bool CheckStatusInstructionReport()
        {
            DateTime dt_Date = DateTime.Today;
            //DateTime dt_Date = Convert.ToDateTime("10/06/2016");
            string strSQL = "select t.*" +
                            " from tas.OIL_LOAD_HEADERS t " +
                            " where to_date(t.UPDATE_DATE) = to_date('" + dt_Date.ToString("dd/MM/yyyy") + "', 'dd/MM/yyyy')" +
                            " and t.LOAD_STATUS = 13" +
                           // " and t.PRINT_AUTO = 1" +
                            " and t.cancel_status = 0 " +
                            " and t.PRINT_STATUS = 0" +
                            " order by t.UPDATE_DATE";
            DataSet vDataset = null;
            DataTable dt;
            bool vRet = false;
            try
            {
                if (fMain.OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
                {
                    dt = vDataset.Tables["TableName"];
                    if (dt != null && dt.Rows.Count != 0)
                    {
                        LoadHeader_NO = dt.Rows[0]["LOAD_HEADER_NO"].ToString();
                        PrintStatus = dt.Rows[0]["PRINT_STATUS"].ToString();
                        vRet = true;
                    }
                }
                
            }
            catch (Exception exp)
            { 
                RaiseEvents("[Print Instruction, CheckStatusLoading] " + exp.Message);
            }
            vDataset = null;
            dt = null;
            return vRet;
        }
        private bool LoadInstructionReport()
        {
            bool vRet = false;
            //if (CheckPrinterOnline(out_PrinterName))
            //    {
                    string strSQL = "select t.*" +
                                    " from rpt.VIEW_LOADING_INTRODUCTION_TM t " +
                                    " where t.LOAD_HEADER_NO=" + LoadHeader_NO;

                    DataSet vDataset = null;
                           vDataset = new DataSet();
                    DataTable dt;
                    try
                    {
                        if (fMain.OraDb.OpenDyns(strSQL, "VIEW_LOADING_INTRODUCTION_TM", ref vDataset))
                        {
                            dt = vDataset.Tables["VIEW_LOADING_INTRODUCTION_TM"];
                            ReportDocument cr = new ReportDocument();
                            if (dt != null && dt.Rows.Count != 0)
                            {
                                string reportPath = fMain.App_Path +"\\Report File\\" + out_ReportPath;
                                if (!File.Exists(reportPath))
                                {
                                    fMain.AddListBox = "The specified report does not exist \n";
                                }
                               
                                cr.Load(reportPath);
                                //cr.SetDataSource(dt);
                                cr.Database.Tables["VIEW_LOADING_INTRODUCTION_TM"].SetDataSource((DataTable)dt);
                                //cr.SetParameterValue(out_Parametername, LoadHeader_NO);
                                cr.SetParameterValue(out_Parametername, LoadHeader_NO);
                                cr.PrintOptions.PrinterName = out_PrinterName;
                                cr.PrintToPrinter(1, false, 0, 0);
                                cr.Dispose();
                                Thread.Sleep(100);
                              
                                RaiseEvents("PRINT INSTRUCTION REPORT LOADING NUMBER: " + LoadHeader_NO);
                                vRet = true;
                            }
                        }
                        Thread.Sleep(100);
                    }
                    catch (Exception exp)
                    { 
                    fMain.LogFile.WriteErrLog("[Print Instruction, Load data report] " + exp.Message);
                    PrintInstructionStepProcess = _PrintInstructionStepProcess.InitialReport;
                    }
                    vDataset = null;
                    dt = null;
               // }
            return vRet;
        }

        void UpdateInstructionToDatabase(string p_LH_NO)
        {
            string vSQL = "update tas.OIL_LOAD_HEADERS t set ";
                   vSQL += " t.PRINT_STATUS=1 where t.LOAD_HEADER_NO='" + p_LH_NO + "'";
            fMain.OraDb.ExecuteSQL(vSQL);
            RaiseEvents("UPDATE PRINT = 1 STATUS of LOAD HEAD NO: " + p_LH_NO);
        }

        private bool GetConfigReport()
        {
            out_ReportID = "52010062";
            //out_ReportID = "52010029";
            string strSQL = "select t.*, rt.*" +
                                    " from tas.VIEW_REPORT_PARA_CONFIG t, tas.PRINTER_TAS rt " +
                                    " where t.PRINTER_ID= rt.PRINTER_ID" +
                                    " and t.Report_ID= " + out_ReportID;
            DataSet vDataset = null;
            DataTable dt;
            bool vRet = false;
            try
            {
                if (fMain.OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
                {
                    dt = vDataset.Tables["TableName"];
                    if (dt != null && dt.Rows.Count != 0)
                    {
                        out_Parametername = dt.Rows[0]["PARAMETER_NAME"].ToString();
                        out_ReportPath = dt.Rows[0]["REPORT_PATH"].ToString();
                        out_PrinterID = dt.Rows[0]["PRINTER_ID"].ToString();
                        out_PrinterName = dt.Rows[0]["PRINTER_NAME"].ToString();
                        vRet = true;
                    }
                }   
            }
            catch (Exception exp)
            { 
                RaiseEvents("[Print Instruction, Get config report] " + exp.Message);
            }
            vDataset = null;
            dt = null;
            return vRet; 
        }

        public bool CheckPrinterOnline(string printerToCheck)
        {
            // Set management scope
            ManagementScope scope = new ManagementScope(@"\root\cimv2");
            scope.Connect();

            // Select Printers from WMI Object Collections
            string query = "SELECT * FROM Win32_Printer WHERE Name='" + printerToCheck + "'";
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);


            bool IsReady = false;
            string printerName = "";
            foreach (ManagementObject printer in searcher.Get())
            {
                printerName = printer["Name"].ToString();
                if (string.Equals(printerName, printerToCheck, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Printer = " + printer["Name"]);
                    if (printer["WorkOffline"].ToString().ToLower().Equals("true"))
                    {
                        // printer is offline by user
                        RaiseEvents("Your Plug-N-Play printer is not connected.");
                    }
                    else
                    {
                        // printer is online
                        IsReady = true;
                        break;
                    }
                }
            }
            return IsReady;
        }
        #endregion
    }
}
