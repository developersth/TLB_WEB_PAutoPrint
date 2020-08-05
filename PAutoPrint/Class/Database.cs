using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.ManagedDataAccess.Client;
using System.Threading;
using System.Data;
using System.Data.OleDb;
using System.Configuration;

namespace PAutoPrintReport
{
    public delegate void EventHandler();


    public class Database : IDisposable
    {
        frmMain fMain;
        //private String mPathIni = "D:\\SAKCTAS\\SAKCTASConfig.ini";     //n/a
        private IniLib.CINI mIni;

        #region Enum Database
        public enum DB_TYPE
        {
            DB_None = -1,
            DB_MASTER = 0,
            DB_BACKUP = 1
        }

        public enum _OracleDbDirection
        {
            OraInput,
            OraOutput
        }

        public enum _OracleDbType
        {
            OraVarchar2,
            OraInt16,
            OraInt32,
            OraInt64,
            OraDate,
            OraLong,
            OraDouble,
            OraSingle,
            OraByte,
            OraDecimal,
            OraBlob
        }
        #endregion

        public sealed class _ParamMember
        {
            public string Name;
            public object Value;
            public Int32 Size;
            public _OracleDbDirection Direction;
            public _OracleDbType DbType;
        }

        //public sealed OracleParameter[] OParamMember;
        //_ParamMember[] _OraParam;
        DB_TYPE currentDB;
        bool isConnectedDB;
        bool isMasterDB;

        int count;
        bool shutdown;
        //string mCnnStrMaster = "User Id=tas;Password=gtas;Data Source=LLTLB";
        string connStrMaster = ConfigurationManager.ConnectionStrings["connStrMaster"].ToString().Trim();
        private OracleConnection oraConn;
        bool oraConnect;
        string oraServiceName = "N/A";
        DateTime oraConnectedDate;
        Thread thrOracle;
        ScanDatabase[] thrScanDB;
        bool thrRunning;
        bool thrRunn;
        Logfile logFile;

        int totalParam;

        public OracleDataReader OraDysReader;

        #region construct and deconstruct
        private bool IsDisposed = false;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
            thrRunn = false;
            mIni = null;
        }
        protected void Dispose(bool Diposing)
        {
            if (!IsDisposed)
            {
                if (Diposing)
                {
                    //Clean Up managed resources
                    Close();
                    mIni = null;
                    thrRunn = false;
                    logFile = null;
                    //PLog = null;

                    //fMain = null;
                }
                //Clean up unmanaged resources
                thrRunn = false;
                thrOracle.Abort();
            }
            IsDisposed = true;
        }
        public Database(frmMain f)
        {
            //fMain = new frmMain();
            fMain = f;
            //mConnOracle = new OracleConnection();
            logFile = new Logfile();
            mIni = new IniLib.CINI();
            currentDB = DB_TYPE.DB_None;
            //ScanDatabase();
            StartThread();

        }
        public Database()
        { }
        ~Database()
        {
            //Dispose(false);
            //mConnOracle = null;

            //Close();
            //PLog = null;
            //fMain = null; 
        }
        #endregion

        public bool ConnectStatus()
        {
            return oraConnect;
        }

        public string ConnectServiceName()
        {
            //string ret = "";
            //switch (CurrentDB)
            //{
            //    case DB_TYPE.DB_None:
            //        ret = "Connect None";
            //        break;
            //    case DB_TYPE.DB_MASTER:
            //        ret = "Connect MASTER";
            //        break;
            //    case DB_TYPE.DB_SUBMASTER:
            //        ret = "Connect BACKUP";
            //        break;
            //}
            //return ret;
            switch (currentDB)
            {
                case DB_TYPE.DB_MASTER:
                    //vMsg = ">Database Connect-> Master";
                    oraServiceName = "Server A";
                    oraServiceName = oraServiceName + "-" + oraConn.DatabaseName + "";
                    break;
                default:
                    //vMsg = ">Database Connect-> NONE";
                    oraServiceName = "Server = N/A";
                    break;
            }
            return oraServiceName.ToUpper();
        }

        private void StartThread()
        {
            thrRunn = true;
            try
            {
                if (thrRunning)
                {
                    return;
                }
                //if (mThread != null)
                //    mThread = null;

                thrOracle = new Thread(this.RunProcess);
                thrRunning = true;
                thrOracle.Name = "thrOracle";
                thrOracle.Start();
            }
            catch (Exception exp)
            {
                thrRunning = false;
            }
        }

        private void RunProcess()
        {
            //Thread.Sleep(500);
            while (true)
            {
                //if (currentDB == DB_TYPE.DB_None)
                //{
                //    currentDB = SelectServer();
                //}
                ScanActiveDatabase();
                if (currentDB != DB_TYPE.DB_None)
                {
                    //Reconnect();
                    //if (mConnect || mShutdown)
                    if (oraConnect || !thrRunn)
                    {
                        break;
                    }
                    // Reconnect();
                }
                System.Threading.Thread.Sleep(5000);
            }
            Thread.Sleep(1000);
            thrRunning = false;
        }

        private void Addlistbox(string pMsg)
        {
            try
            {
                fMain.AddListBox = (object)DateTime.Now + "> " + pMsg;
            }
            catch (Exception exp)
            { }
        }

        void Reconnect()
        {
            if (!oraConnect)
            {
                try
                {
                    oraConn.Close();
                }
                catch (Exception exp)
                { }
                Connect();
            }
        }

        bool Connect()
        {
            string vMsg;
            //mConnOracle = new OracleConnection(mCnnStrMaster);
            switch (currentDB)
            {
                case DB_TYPE.DB_MASTER:
                    //vMsg = ">Database Connect-> Master";
                    oraConn = new OracleConnection(connStrMaster);
                    oraServiceName = "Server A";
                    break;
                default:
                    //vMsg = ">Database Connect-> NONE";
                    oraServiceName = "N/A";
                    break;
            }
            //Addlistbox(vMsg);
            try
            {
                if (currentDB != DB_TYPE.DB_None)
                {
                    oraConn.Open();
                    oraConnect = true;
                    //logFile = new CLogfile();
                    Addlistbox(ConnectServiceName());
                    Addlistbox("Database connect successful.");
                    oraConnectedDate = DateTime.Now;
                    //oraServiceName = oraServiceName + "-" + oraConn.DatabaseName + "";
                    return true;
                }
                else
                {
                    logFile.WriteErrLog("[System]" + "Can not detect Master and Backup Database.");
                    return false;
                }
            }
            catch (Exception exp)
            {
                oraConnect = false;
                thrRunn = false;
                Addlistbox(exp.ToString());
                StartThread();
                return false;
            }
        }

        public void Close()
        {
            shutdown = true;
            Thread.Sleep(500);
            if (oraConn != null)
            {
                try
                {
                    oraConn.Close();
                    oraConn = null;
                    //fMercury = null;
                    logFile = null;
                    //thrScanDB[0].Dispose();
                    //thrScanDB[1].Dispose();
                }
                catch (Exception exp) { }
            }
        }

        void CheckExecute(bool pExe)
        {
            if (pExe)
            {
                count = 0;
                if (!oraConnect)
                    oraConnect = true;
            }
            else
            {
                count += 1;
                if (count >= 3)
                {
                    if (oraConnect)
                    {
                        oraConnect = false;
                        StartThread();
                    }
                    count = 0;
                }
            }
        }

        public bool OpenDyns(string pStrSQL, string pTableName, ref DataSet pDataSet)
        {
            OracleDataAdapter oda;
            DataSet ds = new DataSet();
            bool vCheck = false;
            if (oraConnect)
            {
                try
                {
                    oda = new OracleDataAdapter(pStrSQL, oraConn);
                    oda.Fill(ds, pTableName);
                    pDataSet = ds;
                    CheckExecute(true);
                    vCheck = true;
                }
                catch (Exception exp)
                {
                    CheckExecute(vCheck);
                    Logfile p = new Logfile();
                    logFile.WriteErrLog("[OpenDyns]" + pStrSQL);
                    logFile.WriteErrLog(exp.Message);

                }
                ds = null;
                oda = null;
            }
            return vCheck;
        }

        public bool OpenDyns(string pStrSQL, int pMaxRecord, string pTableName, ref DataSet pDataSet)
        {
            OracleDataAdapter oda;
            DataSet ds = new DataSet();
            bool bCheck = false;
            if (oraConnect)
            {
                try
                {
                    oda = new OracleDataAdapter(pStrSQL, oraConn);
                    oda.Fill(ds, 0, pMaxRecord, pTableName);
                    pDataSet = ds;
                    CheckExecute(true);
                    bCheck = true;
                }
                catch (Exception exp)
                {
                    bCheck = false;
                    CheckExecute(bCheck);
                    logFile.WriteErrLog("[OpenDyns]" + pStrSQL);
                    logFile.WriteErrLog("[OpenDyns]" + exp.Message);
                }
                ds = null;
                oda = null;
            }
            return bCheck;
        }

        public bool OpenOraDataReader(string pStrSQL, ref OracleDataReader pOraDataReader)
        {
            OracleCommand oraCommand;
            bool bCheck = false;
            
            if (oraConnect)
            {
                try
                {
                    oraCommand = new OracleCommand(pStrSQL, oraConn);
                    oraCommand.CommandTimeout = 3;
                    pOraDataReader = oraCommand.ExecuteReader();
                    bCheck = true;
                }
                catch (Exception exp)
                {
                    CheckExecute(false);
                    logFile.WriteErrLog("[ExecuteSQL]" + pStrSQL);
                    logFile.WriteErrLog("[ExecuteSQL]" + exp.Message);
                }
            }
            return bCheck;
        }

        public bool ExecuteSQL(string pStrSQL)
        {
            OracleCommand oCommand;
            bool bCheck = false;

            if (oraConnect)
            {
                try
                {
                    bCheck = true;
                    oCommand = new OracleCommand(pStrSQL, oraConn);
                    oCommand.ExecuteNonQuery();
                    CheckExecute(true);
                }
                catch (Exception exp)
                {
                    bCheck = false;
                    CheckExecute(false);
                    logFile.WriteErrLog("[ExecuteSQL]" + pStrSQL);
                    logFile.WriteErrLog("[ExecuteSQL]" + exp.Message);
                }
            }
            return bCheck;
        }

        public bool ExecuteSQL_PROC(string pStrSQL, COracleParameter pParam)
        {
            OracleCommand vOraCmd;
            bool vCheck = false;
            int vParamNo;
            if (oraConnect)
            {
                try
                {
                    vOraCmd = new OracleCommand();
                    if (pParam == null)
                        return vCheck;

                    foreach (OracleParameter p in pParam.OraParam)
                    {
                        vOraCmd.Parameters.Add(p);
                    }
                    vOraCmd.CommandText = pStrSQL;
                    vOraCmd.CommandType = CommandType.StoredProcedure;
                    vOraCmd.Connection = oraConn;
                    vOraCmd.ExecuteNonQuery();

                    vCheck = true;
                    CheckExecute(true);
                }
                catch (Exception exp)
                {
                    vCheck = false;
                    CheckExecute(false);
                    logFile.WriteErrLog("[ExecuteSQL]" + pStrSQL);
                    logFile.WriteErrLog("[ExecuteSQL]" + exp.Message);
                }
            }
            return vCheck;
        }

        public bool ExecuteSQL(string pStrSQL, COracleParameter pParam)
        {
            OracleCommand vOraCmd;
            bool vCheck = false;
            int vParamNo;
            if (oraConnect)
            {
                try
                {
                    vOraCmd = new OracleCommand();
                    if (pParam == null)
                        return vCheck;

                    foreach (OracleParameter p in pParam.OraParam)
                    {
                        vOraCmd.Parameters.Add(p);
                    }
                    vOraCmd.CommandText = pStrSQL;
                    vOraCmd.CommandType = CommandType.Text;
                    vOraCmd.Connection = oraConn;
                    vOraCmd.ExecuteNonQuery();

                    vCheck = true;
                    CheckExecute(true);
                }
                catch (Exception exp)
                {
                    vCheck = false;
                    CheckExecute(false);
                    logFile.WriteErrLog("[ExecuteSQL]" + pStrSQL);
                    logFile.WriteErrLog("[ExecuteSQL]" + exp.Message);
                }
            }
            return vCheck;
        }

        void GetOracleDbType(Database._OracleDbType pOracleDbType, ref Oracle.ManagedDataAccess.Client.OracleDbType pDAOracleDbType)
        {
            if (pOracleDbType == _OracleDbType.OraByte)
                pDAOracleDbType = OracleDbType.Byte;
            else if (pOracleDbType == _OracleDbType.OraBlob)
                pDAOracleDbType = OracleDbType.Blob;
            else if (pOracleDbType == _OracleDbType.OraDate)
                pDAOracleDbType = OracleDbType.Date;
            else if (pOracleDbType == _OracleDbType.OraDecimal)
                pDAOracleDbType = OracleDbType.Decimal;
            else if (pOracleDbType == _OracleDbType.OraDouble)
                pDAOracleDbType = OracleDbType.Double;
            else if (pOracleDbType == _OracleDbType.OraInt16)
                pDAOracleDbType = OracleDbType.Int16;
            else if (pOracleDbType == _OracleDbType.OraInt32)
                pDAOracleDbType = OracleDbType.Int32;
            else if (pOracleDbType == _OracleDbType.OraInt64)
                pDAOracleDbType = OracleDbType.Int64;
            else if (pOracleDbType == _OracleDbType.OraLong)
                pDAOracleDbType = OracleDbType.Long;
            else if (pOracleDbType == _OracleDbType.OraSingle)
                pDAOracleDbType = OracleDbType.Single;
            else if (pOracleDbType == _OracleDbType.OraVarchar2)
                pDAOracleDbType = OracleDbType.Varchar2;
        }

        void GetOracleDbDirection(_OracleDbDirection pOracleDbDirection, ref System.Data.ParameterDirection pDAOracleDbDirection)
        {
            if (pOracleDbDirection == _OracleDbDirection.OraInput)
                pDAOracleDbDirection = ParameterDirection.Input;
            else if (pOracleDbDirection == _OracleDbDirection.OraOutput)
                pDAOracleDbDirection = ParameterDirection.Output;
        }

        #region "Change active Database Server"
        public void ScanActiveDatabase()
        {
            while (!shutdown)
            {
                //if (mConnect || !mRunn)
                //{
                //    break;
                //}
                DB_TYPE NewDB = SelectServer();
                if (currentDB != NewDB)
                {
                    oraConnect = false;
                    currentDB = NewDB;
                    if (currentDB != DB_TYPE.DB_None)
                    {
                        Reconnect();
                    }
                    else
                    {
                        Addlistbox("Database disconnect.");
                    }
                }
                else
                {
                    if (currentDB != DB_TYPE.DB_None)
                    {
                        if (!oraConnect)
                        {
                            //currentDB = DB_TYPE.DB_None;
                            Addlistbox("Database disconnect.");
                            oraConnect = false;
                            Reconnect();
                        }
                    }
                }
                Thread.Sleep(5000);
            }
        }

        DB_TYPE SelectServer()
        {
            //string ret = mIni.INIRead(mPathIni, "SELECT", "SERVER", "");
            //string ret = "0";
            //return DB_TYPE.DB_MASTER;
            //switch (Convert.ToInt16(ret))
            //{
            //    case 0:
            //        return DB_TYPE.DB_MASTER;
            //    case 1:
            //        return DB_TYPE.DB_SUBMASTER;
            //}
            //return DB_TYPE.DB_None;
            DB_TYPE ret = DB_TYPE.DB_MASTER;
            return ret;
        }

        bool GetConnectServer(DB_TYPE pServer)
        {
            string ret = "0";
            return true;
            switch (pServer)
            {
                case DB_TYPE.DB_None:
                    ret = "0";
                    break;
                case DB_TYPE.DB_MASTER:
                    //ret = mIni.INIRead(mPathIni, "MASTER", "CONNECT", "");
                    break;
                case DB_TYPE.DB_BACKUP:
                    //ret = mIni.INIRead(mPathIni, "SUBMASTER", "CONNECT", "");
                    break;
            }
            return Convert.ToBoolean(Convert.ToInt16(ret));
        }

        bool GetIsMaster(DB_TYPE pServer)
        {
            string ret = "0";
            return true;
            switch (pServer)
            {
                case DB_TYPE.DB_None:
                    ret = "0";
                    break;
                case DB_TYPE.DB_MASTER:
                    //ret = mIni.INIRead(mPathIni, "MASTER", "ISMASTER", "");
                    break;
                case DB_TYPE.DB_BACKUP:
                    //ret = mIni.INIRead(mPathIni, "SUBMASTER", "ISMASTER", "");
                    break;
            }
            return Convert.ToBoolean(Convert.ToInt16(ret));
        }
        #endregion

        class ScanDatabase
        {
            #region construct and deconstruct
            private bool IsDisposed = false;
            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected void Dispose(bool Diposing)
            {
                if (!IsDisposed)
                {
                    if (Diposing)
                    {
                        //Clean Up managed resources
                        //thrScanDB.Abort();
                        thrShutDown = true;
                    }
                    //Clean up unmanaged resources
                }
                IsDisposed = true;
            }

            public ScanDatabase(string pUser, string pPwd, string pDbName)
            {
                dbUser = pUser;
                dbPwd = pPwd;
                dbName = pDbName;
                StartThread();
            }

            ~ScanDatabase()
            {
            }
            #endregion

            Thread thrScanDB;
            string dbUser;
            string dbPwd;
            string dbName;
            OleDbConnection oleConn;

            bool thrShutDown = false;

            void StartThread()
            {
                thrScanDB = new Thread(this.RunProcess);
                thrScanDB.Start();
            }

            void StopThread()
            {
                thrShutDown = true;
            }

            void RunProcess()
            {
                while (!thrShutDown)
                {
                    oleConn = new OleDbConnection();
                    OleDbDataReader oleDataReader = null;
                    try
                    {
                        oleConn.ConnectionString = "Provider=OraOLEDB.Oracle;User ID=" + dbUser +
                                                ";Password=" + dbPwd + ";Data Source=" + dbName + ";";
                        oleConn.Open();
                        string strSQL = "select t.config_data,t.update_date from tas.tas_config t where t.config_id=90"; //check database 1=master active
                        if (OpenDys(strSQL, ref oleDataReader))
                        {
                            if (oleDataReader.HasRows)
                            {
                                oleDataReader.Read();
                                _MasterDataBase = Convert.ToInt32(oleDataReader.GetString(0).ToString());
                                _UpdateDate = oleDataReader.GetDateTime(1);

                                oleDataReader.Close();
                            }
                        }
                        else
                        {
                            _MasterDataBase = -1;
                        }
                        Thread.Sleep(3000);
                    }
                    catch (Exception exp)
                    {
                        _MasterDataBase = -1;
                    }
                    oleConn.Close();
                }
            }

            bool OpenDys(string pStrSQL, ref OleDbDataReader pDataReader)
            {
                OleDbCommand oleCommand ;
                bool vCheck = false;

                if (oleConn.State == ConnectionState.Open)
                {
                    try
                    {
                        oleCommand = new OleDbCommand(pStrSQL, oleConn);
                        oleCommand.CommandTimeout = 1;
                        pDataReader = oleCommand.ExecuteReader();
                        vCheck = true;
                    }
                    catch (Exception exp)
                    {
                        Logfile logFile = new Logfile();
                        logFile.WriteErrLog("[ExecuteSQL]" + pStrSQL);
                        logFile.WriteErrLog("[ExecuteSQL]" + exp.Message);
                        logFile = null;
                    }
                }
                return vCheck;
            }

            #region "Property"
            int _MasterDataBase = -1;
            public int MasterDatabase
            {
                get { return _MasterDataBase; }
            }

            DateTime _UpdateDate;
            public DateTime UpdateDate
            {
                get { return _UpdateDate; }
            }
            #endregion
        }
    }
    public class COracleParameter
    {
        #region "Oracle Parameter"
        public Oracle.ManagedDataAccess.Client.OracleParameter[] OraParam;
        public ParameterDirection OraDirection;
        public OracleDbType OraDbType;
        public OracleParameter mParm;

        public void CreateParameter(int pLength)
        {
            OraParam = new OracleParameter[pLength];
            for (int i = 0; i < OraParam.Length; i++)
            {
                OraParam[i] = new OracleParameter();
            }
        }

        public void CreateOracleParameter(int pLength)
        {
            if (pLength > 0)
            {
                OraParam = new OracleParameter[pLength];
                for (int i = 0; i <= pLength - 1; i++)
                {
                    OraParam[i] = new OracleParameter();
                }
            }
        }

        public void RemoveOracleParameter()
        {
            try
            {
                for (int i = 0; i <= OraParam.Length; i++)
                {
                    OraParam[i].Dispose();
                }
            }
            catch (Exception exp)
            { }
        }

        public void AddOracleParameter(int pIndex, string pName, OracleDbType pDbType, ParameterDirection pDbDirection, int pSize)
        {
            if (pIndex < OraParam.Length)
            {
                OraParam[pIndex].ParameterName = pName;
                OraParam[pIndex].OracleDbType = pDbType;
                OraParam[pIndex].Size = pSize;
                OraParam[pIndex].Direction = pDbDirection;
            }
        }

        public void AddOracleParameter(int pIndex, string pName, OracleDbType pDbType, ParameterDirection pDbDirection)
        {
            if (pIndex < OraParam.Length)
            {
                OraParam[pIndex].ParameterName = pName;
                OraParam[pIndex].OracleDbType = pDbType;
                OraParam[pIndex].Direction = pDbDirection;
                OraParam[pIndex].Size = 512;
            }
        }

        public void AddOracleParameter(int pIndex, string pName, OracleDbType pDbType, int pSize)
        {
            if (pIndex < OraParam.Length)
            {
                OraParam[pIndex].ParameterName = pName;
                OraParam[pIndex].OracleDbType = pDbType;
                OraParam[pIndex].Direction = ParameterDirection.Output;
                OraParam[pIndex].Size = pSize;
            }
        }

        public void AddOracleParameter(int pIndex, string pName, OracleDbType pDbType)
        {
            if (pIndex < OraParam.Length)
            {
                OraParam[pIndex].ParameterName = pName;
                OraParam[pIndex].OracleDbType = pDbType;
                OraParam[pIndex].Direction = ParameterDirection.Output;
                OraParam[pIndex].Size = 512;
            }
        }

        public void RemoveParameter()
        {
            for (int i = 0; i < OraParam.Length - 1; i++)
            {
                OraParam[i].Dispose();
            }
            OraParam = null;
        }

        public void SetParameterValue(int pIndex, object pValue)
        {
            OraParam[pIndex].Value = pValue;
        }

        public void SetOracleParameterValue(string pName, ref OracleParameter pValue)
        {
            for (int i = 0; i < OraParam.Length - 1; i++)
            {
                if (OraParam[i].ParameterName == pName)
                {
                    OraParam[i].Value = pValue;
                    break;
                }
            }
        }

        public void GetOracleParameterValue(int pIndex, ref OracleParameter pParam)
        {
            pParam = OraParam[pIndex];
        }

        public void GetOracleParameterValue(string pName, ref OracleParameter pParam)
        {
            foreach (OracleParameter p in OraParam)
            {
                if (p.ParameterName == pName)
                {
                    pParam = p;
                    return;
                }
            }
        }

        public OracleParameter GetParameter(int pIndex)
        {
            return OraParam[pIndex];
        }

        public OracleParameter GetParameter(string pName)
        {
            foreach (OracleParameter p in OraParam)
            {
                if (p.ParameterName == pName)
                {
                    return p;
                }
            }
            return null;
        }

        #endregion
    }

}
