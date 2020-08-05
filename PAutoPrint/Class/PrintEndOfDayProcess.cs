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
    public class PrintEndOfDayProcess : IDisposable
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
        private string out_ReportID;
        private string[] out_ReportPath;
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

        enum _PrintEndOfDayStepProcess : int
        {
            InitialReport = 0,
            CheckStatusReport = 1,
            LoadReport = 10,
            UpdateStatus = 11,
            ChangeProcess = 30
        }
        #endregion

        frmMain fMain;
        _PrintEndOfDayStepProcess PrintEndOfDayStepProcess;

        int processId = 1;
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

        public PrintEndOfDayProcess(frmMain pFrm)
        {
            fMain = pFrm;
        }

        public PrintEndOfDayProcess(frmMain pFrm, int pId)
        {
            fMain = pFrm;
            processId = pId;
        }

        ~PrintEndOfDayProcess()
        { }
        #endregion

        #region Class Events
        public delegate void ATGEventsHaneler(object pSender, string pEventMsg);
        string logFileName;

        void RaiseEvents(string pSender, string pMsg)
        {
            string vMsg = DateTime.Now + ">[" + pSender + "]" + pMsg;
            logFileName = "Auto Print End OF Day LLTLB";
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
            logFileName = "Auto Print End OF Day LLTLB";
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
            thrMain.Name = processId.ToString() + "thrPrintEndOfDay";
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
            PrintEndOfDayStepProcess = _PrintEndOfDayStepProcess.InitialReport;
            chkResponse = DateTime.Now;
            while (thrRunning)
            {
                try
                {
                    if (thrShutdown)
                        return;
                    switch (PrintEndOfDayStepProcess)
                    {
                        case _PrintEndOfDayStepProcess.InitialReport:
                            if (GetConfigReport())
                            {
                                PrintEndOfDayStepProcess = _PrintEndOfDayStepProcess.CheckStatusReport;
                            }
                            break;
                        case _PrintEndOfDayStepProcess.CheckStatusReport:
                            if (CheckStatusEndOfDayReport())
                            {
                                PrintEndOfDayStepProcess = _PrintEndOfDayStepProcess.LoadReport;
                            }
                            break;
                        case _PrintEndOfDayStepProcess.LoadReport:
                            if (LoadEndOfDayReport())
                                PrintEndOfDayStepProcess = _PrintEndOfDayStepProcess.UpdateStatus;
                            break;
                        case _PrintEndOfDayStepProcess.UpdateStatus:
                            UpdateEndOfDayToDatabase(DateTime.Now);
                            PrintEndOfDayStepProcess = _PrintEndOfDayStepProcess.ChangeProcess; ;
                            break;
                        case _PrintEndOfDayStepProcess.ChangeProcess:
                            PrintEndOfDayStepProcess = _PrintEndOfDayStepProcess.CheckStatusReport;
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

        bool CheckStatusEndOfDayReport()
        {
            DateTime dt_Date = DateTime.Today;
            //DateTime dt_Date = Convert.ToDateTime("10/06/2016");
            string strSQL = @"select t.job_no,t.job_type,t.job_name,t.creation_date,t.created_by from report_printer_job t where to_date(t.creation_date) = to_date(sysdate)";
            DataSet vDataset = null;
            DataTable dt;
            bool vRet = false ;
            try
            {
                if (fMain.OraDb.OpenDyns(strSQL, "TableName", ref vDataset))
                {
                    dt = vDataset.Tables["TableName"];
                    if (dt != null && dt.Rows.Count != 0)
                    {
                        //LoadHeader_NO = dt.Rows[0]["LOAD_HEADER_NO"].ToString();
                        vRet = true;
                    }
                }

            }
            catch (Exception exp)
            { fMain.LogFile.WriteErrLog("[Print Instruction, CheckStatusLoading] " + exp.Message); }
            vDataset = null;
            dt = null;
            return vRet;
        }
        private bool LoadEndOfDayReport()
        {
            DataSet vDataset = null;
            DataTable dt;
            string strSQL;
            bool vRet = false;
            DateTime dt_Date = DateTime.Today;

            //if (CheckPrinterOnline(out_PrinterName))
            //{
                if (out_ReportPath.Length > 0)
                {
                    for (int i = 0; i < out_ReportPath.Length; i++)
                    {
                        //Detail Oil Bulk Loading Volume By Product.rpt
                        if (out_ReportPath[i].Equals("Detail Oil Bulk Loading Volume By Product.rpt"))
                        {
                            strSQL = "select t.* from tas.view_loading_volume_by_product t" +
                                            " where t.eod_date =to_date('" + dt_Date.ToString ("dd/MM/yyyy") + "','dd/MM/yyyy')";

                            try
                            {

                                if (fMain.OraDb.OpenDyns(strSQL, "view_load_volume_report_daily", ref vDataset))
                                {
                                    dt = vDataset.Tables["view_load_volume_report_daily"];
                                    if (dt != null && dt.Rows.Count != 0)
                                    {
                                        ReportDocument cr = new ReportDocument();
                                        string[] reportPath = new string[out_ReportPath.Length];
                                        reportPath[i] = fMain.App_Path + "\\Report File\\" + out_ReportPath[i];
                                        if (!File.Exists(reportPath[i]))
                                        {
                                            fMain.AddListBox = "The specified report does not exist \n";
                                        }
                                        else
                                        {
                                            cr.Load(reportPath[i]);
                                            cr.Database.Tables["view_load_volume_report_daily"].SetDataSource((DataTable)dt);
                                            cr.SetParameterValue(out_Parametername, dt_Date.ToString("dd/MM/yyyy"));
                                            cr.PrintOptions.PrinterName = out_PrinterName;
                                            cr.PrintToPrinter(1, false, 0, 0);
                                            cr.Dispose();
                                            Thread.Sleep(1000);
                                            RaiseEvents("PRINT Detail Oil Bulk Loading Volume By Product.rpt REPORT END OF DAY ");
                                            RaiseEvents(reportPath[i]);
                                            vRet = true;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                RaiseEvents("[Print EndOfDay, Load data report] " + ex.Message);
                                vDataset = null;
                                dt = null;
                                throw;

                            }
                        }
                        //Detail Oil Bulk Loading Mass By Product.rpt
                        if (out_ReportPath[i].Equals("Detail Oil Bulk Loading Mass By Product.rpt"))
                        {
                            strSQL = "select t.* from rpt.view_load_mass_report_daily t" +
                                            " where t.eod_date =to_date('" + DateTime.Now.ToString("dd/MM/yyyy") + "','dd/MM/yyyy')";

                            try
                            {

                                if (fMain.OraDb.OpenDyns(strSQL, "rpt.view_load_mass_report_daily", ref vDataset))
                                {
                                    dt = vDataset.Tables["rpt.view_load_mass_report_daily"];
                                    if (dt != null && dt.Rows.Count != 0)
                                    {
                                        ReportDocument cr = new ReportDocument();
                                        string[] reportPath = new string[out_ReportPath.Length];
                                        reportPath[i] = fMain.App_Path + "\\Report File\\" + out_ReportPath[i];
                                        if (!File.Exists(reportPath[i]))
                                        {
                                            fMain.AddListBox = "The specified report does not exist \n";
                                        }
                                        else
                                        {
                                            cr.Load(reportPath[i]);
                                            cr.Database.Tables["rpt.view_load_mass_report_daily"].SetDataSource((DataTable)dt);
                                            cr.SetParameterValue(out_Parametername, DateTime.Now.ToString("dd/MM/yyyy"));
                                            cr.PrintOptions.PrinterName = out_PrinterName;
                                            cr.PrintToPrinter(1, false, 0, 0);
                                            cr.Dispose();
                                            Thread.Sleep(1000);
                                            RaiseEvents("PRINT Detail Oil Bulk Loading Volume By Product.rpt REPORT END OF DAY ");
                                            RaiseEvents(reportPath[i]);
                                            vRet = true;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                RaiseEvents("[Print EndOfDay, Load data report] " + ex.Message);
                                vDataset = null;
                                dt = null;
                                throw;
                                
                            }
                        }

                        //Detail Oil Bulk Loading Volume By Company.rpt
                        if (out_ReportPath[i].Equals("Detail Oil Bulk Loading Volume By Company.rpt"))
                        {
                            strSQL = "select t.* from rpt.view_load_volume_report_daily t" +
                                            " where t.eod_date =to_date('" + DateTime.Now.ToString("dd/MM/yyyy") + "','dd/MM/yyyy')";

                            try
                            {

                                if (fMain.OraDb.OpenDyns(strSQL, "view_load_volume_report_daily", ref vDataset))
                                {
                                    dt = vDataset.Tables["view_load_volume_report_daily"];
                                    if (dt != null && dt.Rows.Count != 0)
                                    {
                                        ReportDocument cr = new ReportDocument();
                                        string[] reportPath = new string[out_ReportPath.Length];
                                        reportPath[i] = fMain.App_Path + "\\Report File\\" + out_ReportPath[i];
                                        if (!File.Exists(reportPath[i]))
                                        {
                                            fMain.AddListBox = "The specified report does not exist \n";
                                        }
                                        else
                                        {
                                            cr.Load(reportPath[i]);
                                            cr.Database.Tables["view_load_volume_report_daily"].SetDataSource((DataTable)dt);
                                            cr.SetParameterValue(out_Parametername, DateTime.Now.ToString("dd/MM/yyyy"));
                                            cr.PrintOptions.PrinterName = out_PrinterName;
                                            cr.PrintToPrinter(1, false, 0, 0);
                                            cr.Dispose();
                                            Thread.Sleep(1000);
                                            RaiseEvents("PRINT Detail Oil Bulk Loading Volume By Company.rpt REPORT END OF DAY");
                                            RaiseEvents(reportPath[i]);
                                            vRet = true;

                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                RaiseEvents("[Print EndOfDay, Load data report] " + ex.Message);
                                vDataset = null;
                                dt = null;
                                throw;
                            }
                        }
                        //Detail Oil Bulk Loading Mass By Cust.rpt
                        if (out_ReportPath[i].Equals("Detail Oil Bulk Loading Mass By Cust.rpt"))
                        {
                            strSQL = "select t.* from rpt.view_load_mass_report_daily t" +
                                            " where t.eod_date =to_date('" + DateTime.Now.ToString("dd/MM/yyyy") + "','dd/MM/yyyy')";

                            try
                            {

                                if (fMain.OraDb.OpenDyns(strSQL, "view_load_mass_report_daily", ref vDataset))
                                {
                                    dt = vDataset.Tables["view_load_mass_report_daily"];
                                    if (dt != null && dt.Rows.Count != 0)
                                    {
                                        ReportDocument cr = new ReportDocument();
                                        string[] reportPath = new string[out_ReportPath.Length];
                                        reportPath[i] = fMain.App_Path + "\\Report File\\" + out_ReportPath[i];
                                        if (!File.Exists(reportPath[i]))
                                        {
                                            fMain.AddListBox = "The specified report does not exist \n";
                                        }
                                        else
                                        {
                                            cr.Load(reportPath[i]);
                                            cr.Database.Tables["view_load_mass_report_daily"].SetDataSource((DataTable)dt);
                                            cr.SetParameterValue(out_Parametername, DateTime.Now.ToString("dd/MM/yyyy"));
                                            cr.PrintOptions.PrinterName = out_PrinterName;
                                            cr.PrintToPrinter(1, false, 0, 0);
                                            cr.Dispose();
                                            Thread.Sleep(1000);
                                            RaiseEvents("PRINT Detail Oil Bulk Loading Mass By Cust.rpt REPORT END OF DAY ");
                                            RaiseEvents(reportPath[i]);
                                            vRet = true;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                RaiseEvents("[Print EndOfDay, Load data report] " + ex.Message);
                                vDataset = null;
                                dt = null;
                                throw;
                            }
                        }
                        //SummaryVolumeLoadReportByCompany.rpt
                        if (out_ReportPath[i].Equals("SummaryVolumeLoadReportByCompany.rpt"))
                        {
                            strSQL = "select t.* from rpt.view_load_volume_report_daily t" +
                                            " where t.eod_date =to_date('" + DateTime.Now.ToString("dd/MM/yyyy") + "','dd/MM/yyyy')";

                            try
                            {

                                if (fMain.OraDb.OpenDyns(strSQL, "view_load_volume_report_daily", ref vDataset))
                                {
                                    dt = vDataset.Tables["view_load_volume_report_daily"];
                                    if (dt != null && dt.Rows.Count != 0)
                                    {
                                        ReportDocument cr = new ReportDocument();
                                        string[] reportPath = new string[out_ReportPath.Length];
                                        reportPath[i] = fMain.App_Path + "\\Report File\\" + out_ReportPath[i];
                                        if (!File.Exists(reportPath[i]))
                                        {
                                            fMain.AddListBox = "The specified report does not exist \n";
                                        }
                                        else
                                        {
                                            cr.Load(reportPath[i]);
                                            cr.Database.Tables["view_load_volume_report_daily"].SetDataSource((DataTable)dt);
                                            cr.SetParameterValue(out_Parametername, DateTime.Now.ToString("dd/MM/yyyy"));
                                            cr.PrintOptions.PrinterName = out_PrinterName;
                                            cr.PrintToPrinter(1, false, 0, 0);
                                            cr.Dispose();
                                            Thread.Sleep(1000);
                                            RaiseEvents("PRINT SummaryVolumeLoadReportByCompany.rpt REPORT END OF DAY ");
                                            RaiseEvents(reportPath[i]);
                                            vRet = true;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                RaiseEvents("[Print EndOfDay, Load data report] " + ex.Message);
                                vDataset = null;
                                dt = null;
                                throw;
                            }
                        }
                        //SummaryMassLoadReportByCompany.rpt
                        if (out_ReportPath[i].Equals("SummaryMassLoadReportByCompany.rpt"))
                        {
                            strSQL = "select t.* from rpt.view_load_mass_report_daily t" +
                                            " where t.eod_date =to_date('" + DateTime.Now.ToString("dd/MM/yyyy") + "','dd/MM/yyyy')";

                            try
                            {

                                if (fMain.OraDb.OpenDyns(strSQL, "view_load_mass_report_daily", ref vDataset))
                                {
                                    dt = vDataset.Tables["view_load_mass_report_daily"];
                                    if (dt != null && dt.Rows.Count != 0)
                                    {
                                        ReportDocument cr = new ReportDocument();
                                        string[] reportPath = new string[out_ReportPath.Length];
                                        reportPath[i] = fMain.App_Path + "\\Report File\\" + out_ReportPath[i];
                                        if (!File.Exists(reportPath[i]))
                                        {
                                            fMain.AddListBox = "The specified report does not exist \n";
                                        }
                                        else
                                        {
                                            cr.Load(reportPath[i]);
                                            cr.Database.Tables["view_load_mass_report_daily"].SetDataSource((DataTable)dt);
                                            cr.SetParameterValue(out_Parametername, DateTime.Now.ToString("dd/MM/yyyy"));
                                            cr.PrintOptions.PrinterName = out_PrinterName;
                                            cr.PrintToPrinter(1, false, 0, 0);
                                            cr.Dispose();
                                            Thread.Sleep(1000);
                                            RaiseEvents("PRINT SummaryMassLoadReportByCompany REPORT END OF DAY ");
                                            RaiseEvents(reportPath[i]);
                                            vRet = true;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                RaiseEvents("[Print EndOfDay, Load data report] " + ex.Message);
                                vDataset = null;
                                dt = null;
                                throw;
                            }
                        }
                    }
                }
                vDataset = null;
                dt = null;
           // }
            return vRet;
        }
    
        void UpdateEndOfDayToDatabase(DateTime p_LH_NO)
        {
            string vDate = p_LH_NO.ToString("dd/MM/yyyy");

            string vSQL = "delete from report_printer_job ";
            fMain.OraDb.ExecuteSQL(vSQL);
            RaiseEvents("ลบข้อมูล report_printer_job : " + p_LH_NO);
        }

        private bool GetConfigReport()
        {

            out_ReportID = "52010010";
            string strSQL = "select t.*, rt.*" +
                                    " from tas.VIEW_REPORT_PARA_CONFIG t, tas.PRINTER_TAS rt " +
                                    " where t.PRINTER_ID= rt.PRINTER_ID" +
                                    " and t.PRINTER_ID= " + out_ReportID;
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
                        out_ReportPath = new string[dt.Rows.Count];
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            out_ReportPath[i] = dt.Rows[i]["REPORT_PATH"].ToString();

                        }
                        out_Parametername = dt.Rows[0]["PARAMETER_NAME"].ToString();
                        out_PrinterID = dt.Rows[0]["PRINTER_ID"].ToString();
                        out_PrinterName = dt.Rows[0]["PRINTER_NAME"].ToString();
                        vRet = true;
                    }
                }
            }
            catch (Exception exp)
            { fMain.LogFile.WriteErrLog("[Print Instruction, Get config report] " + exp.Message); }
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
                        fMain.AddListBox = "Your Plug-N-Play printer is not connected.";
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
