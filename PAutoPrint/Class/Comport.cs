using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Net.Sockets;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;
using System.Net.NetworkInformation;

namespace PAutoPrintReport
{
    public class Comport : IDisposable
    {
        public Action<string, string> DataReceived;
        public Action<string, string> OnCommEvents;

        enum _ComportType : int
        {
            SerialPort = 0
            ,
            TCP_IPPort = 1
        }

        //_ComportType comportType;
        struct _ComportMember
        {
            public bool IsTCP;
            public string PortNo;
            public string PortSetting;
            public _ComportType PortType;
            public int IP_Port;
            public string IP_Address;
            public bool IsOpen;
            public bool IsIdle;
            public bool IsEnable;
            public double ComportID;
        }

        _ComportMember ComportMember;
        frmMain fMain;
        Logfile logFile;
        //cAccuLoads[] cAccu;
        DateTime mResponseTime;
        Thread thrProcessPort;

        bool connect;
        bool thrShutdown;
        bool thrRunning;
        bool thrRunn;
        //private bool istcp;
        //private string port;

        public SerialPort Sp;
        private TcpClient tc;
        Stream stm;
        //public Double ComprtID;

        public string TimeSend;
        public string TimeRecv;
        private string mRecv;
        public bool DiagComport;

        private byte[] mRecbyte = new byte[512];
        private byte[] mSenbyte = new byte[512];
        private string mBay;
        private string mIsland;
        string comportMsg;
        //string mOwnerName;
        bool comportResponse;
        int countWriteLog;
        string mMeterName;
        private static object thrLock = new object();
        private int totalCharacterBit = 11; //default value -> one start, 8 data, 1 parity, 1 stop
        private Single characterTime;

        #region construct and deconstruct
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
                    thrRunn = false;
                    thrShutdown = true;
                    ComportMember.IsIdle = true;
                    if ((thrProcessPort != null) && (thrProcessPort.ThreadState == ThreadState.Running))
                        thrProcessPort.Abort();
                    Thread.Sleep(5);
                    if (IsOpen())
                        ClosePort();

                    logFile = null;
                    tc = null;
                    Sp = null;
                    stm = null;

                }
                //Clean up unmanaged resources
            }
            thrShutdown = true;
            IsDisposed = true;
        }
        
        public Comport(frmMain pFrm, ref System.IO.Ports.SerialPort pPort)
        {
            fMain = pFrm;
            Sp = pPort;
            logFile = new Logfile();
        }

        public Comport(frmMain pFrm, ref System.IO.Ports.SerialPort pPort,double pCompID)
        {
            fMain = pFrm;
            Sp = pPort;
            ComportMember.ComportID = pCompID;
            logFile = new Logfile();
        }
        ~Comport()
        {
            thrShutdown = true;
        }
        #endregion

        #region property
        public double ComportID
        {
            get { return ComportMember.ComportID; }
        }
        #endregion

        #region thread
        public void StartThread()
        {
            thrRunn = true;
            mResponseTime = DateTime.Now;
            thrRunning = false;
            try
            {
                if (thrRunning)
                {
                    return;
                }
                thrProcessPort = new Thread(this.RunProcess);
                thrRunning = true;
                thrProcessPort.Name = "Com_id[" + ComportMember.ComportID + "]";
                thrProcessPort.Start();
            }
            catch (Exception exp)
            {
                thrRunning = false;
            }

        }

        public void StartThread(string pMeterName)
        {
            thrRunn = true;
            mResponseTime = DateTime.Now;
            thrRunning = false;
            mMeterName = pMeterName;

            if (ComportMember.IsTCP)
            {
                comportMsg = "Open comport =" + ComportMember.IP_Address + ":" + ComportMember.IP_Port;
            }
            else
            {
                comportMsg = "Open comport =" + ComportMember.PortNo + ":" + ComportMember.PortSetting;
            }

            comportMsg = comportMsg + (ComportMember.IsEnable ? "" : "[Comport disable!!!]").ToString();
            RaiseEvents(comportMsg);

            try
            {
                if (thrRunning)
                {
                    return;
                }
                thrProcessPort = new Thread(this.RunProcess);
                thrRunning = true;
                thrProcessPort.Name = "Com_id[" + ComportMember.ComportID + "]";
                thrProcessPort.Start();
            }
            catch (Exception exp)
            {
                thrRunning = false;
            }

        }

        private void RunProcess()
        {
            int vCount = 0;
            Thread.Sleep(500);
            while (!thrShutdown)
            {
                if (ComportMember.IsEnable)
                {
                    if (vCount == 0)
                    {
                        OpenPort();
                        vCount++;
                    }
                    else
                    {
                        vCount += 1;
                        if (vCount > 6)
                            vCount = 0;
                    }
                }
                else
                    InitialPort(ComportMember.ComportID);
                //if (mConnect || mShutdown)
                if (connect || !thrRunn)
                    break;
                Thread.Sleep(1000);

            }
            //ClosePort();
            Thread.Sleep(100);
            thrRunning = false;
        }

        #endregion

        public void CheckResponse(bool pResponse)
        {
            if (pResponse)
            {
                mResponseTime = DateTime.Now;
                comportResponse = pResponse;
            }
            else
            {
                DateTime vDateTime = DateTime.Now;
                var vDiff = (vDateTime - mResponseTime).TotalSeconds;

                if ((vDiff > 5) && (comportResponse = true))
                {
                    comportResponse = false;
                    //if (sp.IsOpen)
                    //{
                    //    ClosePort();
                    //    RunProcess();
                    //}
                }
            }

        }

        public void InitialPort(Double pCompID)
        {
            DataSet vDataSet = new DataSet();
            DataTable dt;
            bool b;

            ComportMember.ComportID = pCompID;
            string strSQL = "select t.comport_no,t.comport_setting,t.comport_type" +
                            " from tas.VIEW_ATG_COMPORT t" +
                            " where t.comp_id=" + ComportMember.ComportID;

            if (fMain.OraDb.OpenDyns(strSQL, "TableName", ref vDataSet))
            {
                dt = vDataSet.Tables["TableName"];
                ComportMember.PortNo = dt.Rows[0]["comport_no"].ToString();
                if (!(Convert.IsDBNull(dt.Rows[0]["comport_setting"])))
                    ComportMember.PortSetting = dt.Rows[0]["comport_setting"].ToString();
                ComportMember.PortType = CheckComportType(dt.Rows[0]["comport_type"].ToString());

                ComportMember.IsEnable = Convert.ToBoolean(dt.Rows[0]["enabled"]);
                //if (!(Convert.IsDBNull(dt.Rows[0]["ip_address"])))
                //    argumentPort.IP_Address = dt.Rows[0]["ip_address"].ToString();
            }

            if (ComportMember.PortType == _ComportType.TCP_IPPort)
            {
                ComportMember.IsTCP = true;
                comportMsg = "Open comport =" + ComportMember.IP_Address + ":" + ComportMember.IP_Port;
            }
            else
            {
                ComportMember.IsTCP = false;
                comportMsg = "Open comport =" + ComportMember.PortNo + ":" + ComportMember.PortSetting;
            }

            comportMsg = comportMsg + (ComportMember.IsEnable ? "" : "[Comport disable!!!]").ToString();
            //Addlistbox(mMsg);
            vDataSet = null;
            //OpenPort();
        }

        public void InitialPort()
        {
            DataSet vDataSet = new DataSet();
            DataTable dt;
            bool b;

            string strSQL = "select t.comport_no,t.comport_setting,t.comport_type" +
                            " from tas.VIEW_ATG_COMPORT t" +
                            " where t.comp_id=" + ComportMember.ComportID;

            if (fMain.OraDb.OpenDyns(strSQL, "TableName", ref vDataSet))
            {
                dt = vDataSet.Tables["TableName"];
                ComportMember.PortNo = dt.Rows[0]["comport_no"].ToString();
                if (!(Convert.IsDBNull(dt.Rows[0]["comport_setting"])))
                    ComportMember.PortSetting = dt.Rows[0]["comport_setting"].ToString();
                ComportMember.PortType = CheckComportType(dt.Rows[0]["comport_no"].ToString());

                ComportMember.IsEnable = true;
            }

            if (ComportMember.PortType == _ComportType.TCP_IPPort)
            {
                ComportMember.IsTCP = true;
                comportMsg = "Initial comport =" + ComportMember.IP_Address + ":" + ComportMember.IP_Port;
            }
            else
            {
                ComportMember.IsTCP = false;
                comportMsg = "Initial comport id=" + ComportMember.ComportID + ":" + ComportMember.PortNo + "," + ComportMember.PortSetting;
            }

            comportMsg = comportMsg + (ComportMember.IsEnable ? "" : "[Comport disable!!!]").ToString();
            RaiseEvents(comportMsg);
            vDataSet = null;
            //OpenPort();
        }

        #region Comport & TCP/IP

        private bool OpenPort()
        {
            bool vIsAvailable = true;
            try
            {
                if (ComportMember.IsTCP)
                {
                    comportMsg = "Open comport =" + ComportMember.IP_Address + " : " + ComportMember.IP_Port + ".";
                    Application.DoEvents();
                    IPGlobalProperties ipGlobalProperties = IPGlobalProperties.GetIPGlobalProperties();
                    TcpConnectionInformation[] tcpConnInfoArray = ipGlobalProperties.GetActiveTcpConnections();
                    foreach (TcpConnectionInformation tcpi in tcpConnInfoArray)
                    {
                        if (tcpi.LocalEndPoint.Port == ComportMember.IP_Port)
                        {
                            vIsAvailable = false;
                            break;
                        }
                    }

                    if (vIsAvailable)
                    {
                        tc = new TcpClient(ComportMember.IP_Address, ComportMember.IP_Port);
                        //tc = new TcpClient("192.168.1.193", 7734);
                        tc.SendTimeout = 1000;
                        stm = tc.GetStream();
                        comportMsg = "Comport =" + ComportMember.IP_Address + " : " + ComportMember.IP_Port + " Open successfull.";
                        RaiseEvents(comportMsg);
                        ComportMember.IsOpen = true;
                    }
                }
                else//serial port
                {
                    if (Sp.IsOpen)
                    {
                        //ComportMember.IsOpen = true;
                        comportMsg = "Port is opened by another application!";
                        RaiseEvents(comportMsg);
                        return true;
                    }
                    comportMsg = "Open comport =" + ComportMember.PortNo + ":" + ComportMember.PortSetting + ".";
                    RaiseEvents(comportMsg);
                    string[] spli = ComportMember.PortSetting.Split(',');       //9600,N,8,1 -> baud,parity,data,stop
                    if (spli.Length != 4)
                        return false;
                    //mSp = new SerialPort();
                    //mSp.DataReceived += SerialPort_DataReceived;
                    Sp.PortName = ComportMember.PortNo;
                    Sp.BaudRate = Int32.Parse(spli[0]);
                    switch (spli[1].ToString().ToUpper().Trim())
                    {
                        case "E": Sp.Parity = Parity.Even; break;
                        case "M": Sp.Parity = Parity.Mark; break;
                        case "O": Sp.Parity = Parity.Odd; break;
                        case "S": Sp.Parity = Parity.Space; break;
                        default: Sp.Parity = Parity.None; break;
                    }
                    Sp.DataBits = Int32.Parse(spli[2]);
                    switch (spli[3].Trim())
                    {
                        case "1": Sp.StopBits = StopBits.One; break;
                        case "1.5": Sp.StopBits = StopBits.OnePointFive; break;
                        case "2": Sp.StopBits = StopBits.Two; break;
                        default: Sp.StopBits = StopBits.None; break;
                    }

                    Sp.ReadTimeout = 100;
                    Sp.WriteTimeout = 500;
                    Sp.Open();
                    Sp.DiscardInBuffer();
                    Sp.DiscardOutBuffer();
                }
                ComportMember.IsOpen = true;
                ComportMember.IsIdle = true;
                connect = true;
                comportMsg = "Comport =" + ComportMember.PortNo + ":" + ComportMember.PortSetting + " Open successfull.";
                RaiseEvents(comportMsg);
            }
            catch (Exception exp)
            {
                //Addlistbox(exp.Message.ToString());
                if (countWriteLog > 120)
                {
                    countWriteLog = 0;
                    logFile.WriteErrLog("[Comport]" + exp.ToString());
                    ComportMember.IsOpen = false;
                }
                else
                    countWriteLog += 1;
                return false;
            }
            InitialCharacterTime();
            return true;
        }

        void InitialCharacterTime()
        {
            if (Sp.DataBits == 8)
                totalCharacterBit = 11;
            else
                totalCharacterBit = 10;

            characterTime = Convert.ToSingle((totalCharacterBit * (Single)1000) / (Single)Sp.BaudRate); //in ms
        }

        public int CalculateSilentTime(int pCharacterLen) //return in ms
        {
            //nitialCharacterTime();
            if (pCharacterLen <= 8)
                pCharacterLen = 8;
            return Convert.ToInt16(characterTime * (pCharacterLen + 1) * 8 * 3.5);
        }

        public string GetComportDescription()
        {
            string vMsg = "";
            if (ComportMember.IsTCP)
            {
                vMsg = "Comport:" + ComportMember.IP_Address + ":" + ComportMember.IP_Port;
            }
            else
            {
                vMsg = "Comport:" + ComportMember.PortNo + ":" + ComportMember.PortSetting;
            }
            return vMsg;
        }

        public string GetComportNo()
        {
            return ComportMember.PortNo;
        }

        public void ClosePort()
        {
            try
            {
                if (ComportMember.IsTCP)
                {
                    tc.Close();
                    comportMsg = "Comport =" + ComportMember.IP_Address + " : " + ComportMember.IP_Port +
                        " close successfull.";
                }
                else
                {
                    if (Sp.IsOpen)
                    {
                        Sp.DataReceived -= new SerialDataReceivedEventHandler(SerialPort_DataReceived);
                    }
                    Sp.DiscardInBuffer();
                    Sp.DiscardOutBuffer();
                    Sp.Close();

                    ComportMember.IsOpen = false;
                    comportMsg = "Comport =" + ComportMember.PortNo + " : " + ComportMember.PortSetting +
                            " close successfull.";
                }
                RaiseEvents(comportMsg);
            }

            catch (Exception exp) { /*PLog.WriteErrLog(exp.ToString());*/ }
        }

        public bool IsOpen()
        {
            return ComportMember.IsOpen;
        }

        public bool IsIdle
        {
            get { return ComportMember.IsIdle; }
            set { ComportMember.IsIdle = value; }
        }

        private _ComportType CheckComportType(string pCompNo)
        {
            if (pCompNo.IndexOf("COM") >= 0)
            {
                return _ComportType.SerialPort;
            }
            else
            {

                return _ComportType.TCP_IPPort;
            }
        }
        //private bool TestConnection()
        //{
        //try
        //{
        //    return !(TCPSocket.Poll(1000, SelectMode.SelectRead) && TCPSocket.Available == 0);
        //}
        //catch (SocketException) { return false; }
        //}
        #endregion

        public void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                //no. of data at the port
                int ByteToRead = Sp.BytesToRead;

                //create array to store buffer data
                byte[] inputData = new byte[ByteToRead];

                //read the data and store
                Sp.Read(inputData, 0, ByteToRead);

                string s = Encoding.ASCII.GetString(inputData);
                var copy = DataReceived;
                if (copy != null) copy(s, "");

            }
            catch (SystemException exp)
            {
                logFile.WriteErrLog("[Comport]"+exp.ToString());

            }
        }

        public void SerialPort_DataReceived1(string pOwnerName)
        {
            try
            {
                //no. of data at the port
                int ByteToRead = Sp.BytesToRead;

                //create array to store buffer data
                byte[] inputData = new byte[ByteToRead];

                //read the data and store
                Sp.Read(inputData, 0, ByteToRead);

                string s = Encoding.ASCII.GetString(inputData);
                var copy = DataReceived;
                if (copy != null) copy(s, pOwnerName);

            }
            catch (SystemException ex)
            {
                //MessageBox.Show(ex.Message, "Data Received Event");
            }
        }

        public string ReceiveData()
        {
            string vRecv = "";
            //Thread.Sleep(70);
            lock (thrLock)
            {
                if (IsOpen())
                {
                    try
                    {
                        if (ComportMember.IsTCP)
                        {
                            stm.Read(mRecbyte, 0, mRecbyte.Length);
                            vRecv = ASCIIEncoding.UTF8.GetString(mRecbyte);
                        }
                        else
                        {
                            //mRecv = mSp.ReadExisting();
                            //mSp.DiscardInBuffer();
                            //var copy = DataReceived;
                            //if (copy != null) copy(recv);
                            //no. of data at the port
                            //int ByteToRead = mSp.BytesToRead;
                            int OldByteToRead = -1;
                            int NewByteToRead = 0;
                            int i = 0;
                            while ((OldByteToRead < NewByteToRead) && i <= 20)
                            {
                                Thread.Sleep(10);
                                i++;
                                if (i > 2)
                                {
                                    if (NewByteToRead != Sp.BytesToRead)
                                        NewByteToRead = Sp.BytesToRead;
                                    else
                                        OldByteToRead = NewByteToRead;
                                }
                            }
                            //create array to store buffer data
                            byte[] inputData = new byte[NewByteToRead];

                            //read the data and store
                            Sp.Read(inputData, 0, NewByteToRead);
                            vRecv = Encoding.ASCII.GetString(inputData);
                            Sp.DiscardInBuffer();
                        }
                    }
                    catch (Exception exp)
                    {
                        CheckResponse();
                        vRecv = "";
                    }
                }
                else
                {
                    CheckResponse();
                }
            }
            return vRecv;
        }

        public void ReceiveData(ref byte[] pByteRecv)
        {
            //Thread.Sleep(70);
            pByteRecv = new byte[1];
            if (IsOpen())
            {
                try
                {
                    if (ComportMember.IsTCP)
                    {
                        stm.Read(mRecbyte, 0, mRecbyte.Length);
                        pByteRecv = new byte[mRecbyte.Length];
                        Array.Copy(mRecbyte, pByteRecv, mRecbyte.Length);
                        //vRecv = ASCIIEncoding.UTF8.GetString(mRecbyte);
                    }
                    else
                    {
                        //mRecv = mSp.ReadExisting();
                        //mSp.DiscardInBuffer();
                        //var copy = DataReceived;
                        //if (copy != null) copy(recv);
                        //no. of data at the port
                        int ByteToRead = Sp.BytesToRead;

                        //create array to store buffer data
                        byte[] inputData = new byte[ByteToRead];

                        //read the data and store
                        Sp.Read(inputData, 0, ByteToRead);
                        pByteRecv = new byte[inputData.Length];
                        Array.Copy(inputData, pByteRecv, inputData.Length);
                        //vRecv = Encoding.ASCII.GetString(inputData);
                        Sp.DiscardInBuffer();
                    }
                }
                catch (Exception exp)
                {
                    CheckResponse();
                    //vRecv = "";
                }
            }
            else
            {
                CheckResponse();
            }
        }

        public void SendDataTransaction(string pCmd, ref byte[] pRevcByte)
        {
            byte[] vSend;
            pRevcByte = new byte[1];
            lock (thrLock)
            {
                if (IsOpen())
                {
                    try
                    {
                        TimeSend = DateTime.Now.TimeOfDay.ToString();
                        if (ComportMember.IsTCP)
                        {
                            mSenbyte = new byte[pCmd.Length * sizeof(char)];
                            System.Buffer.BlockCopy(pCmd.ToCharArray(), 0, mSenbyte, 0, mSenbyte.Length);
                            stm.Write(mSenbyte, 0, mSenbyte.Length);
                        }
                        else
                        {
                            Sp.Write(pCmd);
                        }
                        Thread.Sleep(500);
                        TimeRecv = DateTime.Now.TimeOfDay.ToString();

                        ReceiveData(ref pRevcByte);
                    }
                    catch (Exception exp)
                    {
                        CheckResponse();
                    }
                }
            }
        }

        private string SendData(string pMsg)
        {
            byte[] vSend;
            lock (thrLock)
            {
                if (IsOpen())
                {
                    try
                    {
                        //Thread.Sleep(CalculateSilentTime(iMsg.Length));
                        TimeSend = DateTime.Now.TimeOfDay.ToString();
                        if (ComportMember.IsTCP)
                        {
                            mSenbyte = new byte[pMsg.Length * sizeof(char)];
                            System.Buffer.BlockCopy(pMsg.ToCharArray(), 0, mSenbyte, 0, mSenbyte.Length);
                            stm.Write(mSenbyte, 0, mSenbyte.Length);
                        }
                        else
                        {
                            Sp.Write(pMsg);
                        }
                        Thread.Sleep(300);
                        Thread.Sleep(CalculateSilentTime(pMsg.Length));
                        mRecv = ReceiveData();
                        TimeRecv = DateTime.Now.TimeOfDay.ToString();

                        //Thread.Sleep(200);
                        //Thread.Sleep(CalculateSilentTime(mRecv.Length));
                    }
                    catch (Exception exp)
                    {
                        CheckResponse();
                        return "";
                    }
                }
            }
            return mRecv;
        }

        public void SendData(byte[] pMsg)
        {
            //Application.DoEvents();
            byte[] vSend = null;
            lock (thrLock)
            {
                if (IsOpen())
                {
                    try
                    {
                        if (ComportMember.IsTCP)
                        {
                            Array.Copy(pMsg, vSend, pMsg.Length);
                            System.Buffer.BlockCopy(pMsg, 0, vSend, 0, pMsg.Length);
                        }
                        else
                        {
                            Thread.Sleep(100);
                            //mSp.DiscardOutBuffer();
                            string s = Encoding.ASCII.GetString(pMsg);
                            Sp.Write(pMsg, 0, pMsg.Length);

                        }
                    }
                    catch (Exception exp)
                    {
                        CheckResponse();
                    }
                }
            }
        }

        private byte[] BuildMsg(byte pCR_address, string pMsg)
        {
            //string strData;
            byte[] vCR_msg = null;
            //byte[] cr_data = null;

            //CRSendNextBlock(cr_address, ref cr_msg, msg);

            byte mSTX = 2;
            byte R = 82;
            byte mETX = 3;

            byte[] asc = Encoding.ASCII.GetBytes(pMsg);

            byte[] vMsg = new byte[1 + 2 + 1 + asc.Length + 1 + 1 + 1];
            for (int i = 0; i < vMsg.Length; i++)
            {
                switch (i)
                {   //string.Format("{0:X}",(int)(msg[i-1]));
                    case 0:
                        vMsg[i] = mSTX;
                        break;
                    case 1:
                        asc = Encoding.ASCII.GetBytes(pCR_address.ToString());
                        vMsg[i + 1] = asc[0];
                        asc = Encoding.ASCII.GetBytes("0");
                        vMsg[i] = asc[0];
                        break;
                    case 2:
                        break;
                    case 3:
                        vMsg[i] = R;
                        break;
                    default:
                        if (i == vMsg.Length - 3)   //DMY=00
                        {
                            vMsg[i] = (byte)(32);
                        }
                        else if (i == vMsg.Length - 2)
                            break;
                        else if (i == vMsg.Length - 1)
                        {
                            vMsg[i] = mETX;
                        }
                        //else if (i == 4)    
                        //{

                        //    bMsg[i] = ESC;
                        //}
                        else
                            vMsg[i] = (asc[i - 4]);
                        break;
                }
            }
            vCR_msg = vMsg;
            CalCSUM(vCR_msg, ref vCR_msg[vCR_msg.Length - 2]);
            return vCR_msg;
        }

        private void CalCSUM(byte[] pMsg, ref byte pCSUM)
        {
            int vCSUM;
            vCSUM = 0;
            for (int i = 0; i < pMsg.Length - 1; i++)
            {
                vCSUM += pMsg[i];
            }
            vCSUM &= 127;    //bit AND 0x7F
            vCSUM = ~vCSUM;

            pCSUM = (byte)(vCSUM & 127);
            pCSUM += 1;
        }

        private void CheckResponse()
        {
            return;
            DateTime vDateTime = DateTime.Now;
            var vDiff = (vDateTime - mResponseTime).TotalSeconds;

            if ((vDiff > 5) && (comportResponse == true))
            {
                mResponseTime = DateTime.Now;
                comportResponse = false;
                if (ComportMember.IsOpen)
                {
                    ClosePort();
                    StartThread();
                }
            }
        }

        #region class events
        public delegate void ComportEventsHandler(object sender, string message);
        public event ComportEventsHandler OnComportEvents;
        string logFileName;

        void RaiseEvents(string pSender, string pMsg)
        {
            string vMsg = DateTime.Now + ">[" + pSender + "]" + pMsg;
            logFileName = "ATG";
            try
            {
                fMain.AddListBox = "" + logFileName + ">" + vMsg;
                fMain.LogFile.WriteLog(logFileName, vMsg);
            }
            catch (Exception exp)
            { }
        }

        void RaiseEvents(string pMsg)
        {
            string vMsg = DateTime.Now + ">" + pMsg;
            logFileName = "ATG";
            try
            {
                fMain.AddListBox = "" + logFileName + ">" + vMsg;
                //fMain.LogFile.WriteLog(logFileName, vMsg);
            }
            catch (Exception exp)
            { }
        }
        //void RaiseEvents(string pMsg)
        //{
        //    if (OnComportEvents != null)
        //    {
        //        OnComportEvents((string)ComportMember.PortNo, pMsg);
        //    }
        //}
        #endregion
    }
}
