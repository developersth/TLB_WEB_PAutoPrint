using System;
using System.Collections.Generic;
using System.Text;
using System.IO;


namespace PAutoPrintReport
{
    public class Logfile
    {
        string dirLog = Directory.GetCurrentDirectory() + "\\Log\\";
        string dirComportLog = Directory.GetCurrentDirectory() + "\\ComportLog\\";

        #region Constructor / Deconstructor
        public Logfile()
        {
            ScanDeleteLog(dirLog, 120);
            ScanDeleteLog(dirComportLog, 120);
        }
        ~Logfile()
        {
        }
        #endregion

        public void CreateFolderLog()
        {
            if (!Directory.Exists(dirLog))
            {
                Directory.CreateDirectory(dirLog);
            }

            if (!Directory.Exists(dirComportLog))
            {
                Directory.CreateDirectory(dirComportLog);
            }
        }

        public void WriteLog(string pFileName, string pMsg)
        {
            string vFlieName = "Log" + pFileName + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
            string vHeaderText = "" + pFileName + ">";

            CreateFolderLog();
            using (System.IO.StreamWriter pLogFile = new StreamWriter(dirLog + vFlieName, true))
            {
                //pLogFile.WriteLine(vHeaderText + pMsg);
                pLogFile.WriteLine(pMsg);
                pLogFile.Dispose();
            }

        }

        public void WriteLog(string pFileName, string[] pMsg)
        {
            string vFlieName = "Log" + pFileName + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
            string vHeaderText = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "><" + pFileName + ">";

            CreateFolderLog();

            for (int i = 0; i < pMsg.Length; i++)
            {
                pMsg[i] = pMsg[i].PadLeft(pMsg[i].Length + vHeaderText.Length, ' ');    //padding with space character
                //switch(i)
                //{
                //    case 0:
                //        msg[i] = HeaderText + msg[i];
                //        break;
                //    default:
                //        msg[i] = msg[i].PadLeft(msg[i].Length + HeaderText.Length , ' ');    //padding with space character
                //        break;
                //}
            }
            using (System.IO.StreamWriter pLogFile = new StreamWriter(dirLog + vFlieName, true))
            {
                for (int i = 0; i < pMsg.Length; i++)
                {
                    pLogFile.WriteLine(pMsg[i]);
                }
                pLogFile.Dispose();
            }
        }

        public void WriteComportLog(string pFileName, string pMsg)
        {
            string vFlieName = "Log" + pFileName + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
            string vHeaderText = "" + pFileName + ">";

            CreateFolderLog();
            using (System.IO.StreamWriter pLogFile = new StreamWriter(dirComportLog + vFlieName, true))
            {
                //pLogFile.WriteLine(vHeaderText + pMsg);
                pLogFile.WriteLine(pMsg);
                pLogFile.Dispose();
            }

        }

        public void WriteComportLog(string pFileName, string[] pMsg)
        {
            string vFlieName = "Log" + pFileName + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
            string vHeaderText = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "><" + pFileName + ">";

            CreateFolderLog();

            for (int i = 0; i < pMsg.Length; i++)
            {
                //pMsg[i] = pMsg[i].PadLeft(pMsg[i].Length + vHeaderText.Length, ' ');    //padding with space character
                switch (i)
                {
                    case 0:
                        pMsg[i] = pMsg[i];
                        break;
                    default:
                        pMsg[i] = pMsg[i].PadLeft(pMsg[i].Length + vHeaderText.Length, ' ');    //padding with space character
                        break;
                }
            }
            using (System.IO.StreamWriter pLogFile = new StreamWriter(dirComportLog + vFlieName, true))
            {
                for (int i = 0; i < pMsg.Length; i++)
                {
                    pLogFile.WriteLine(pMsg[i]);
                }
                pLogFile.Dispose();
            }
        }

        public void WriteErrLog(string pMsg)
        {
            string vFlieName = "Err_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
            string vHeaderText = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "> ";

            CreateFolderLog();
            using (System.IO.StreamWriter pLogFile = new StreamWriter(dirLog + vFlieName, true))
            {
                pLogFile.WriteLine(vHeaderText + pMsg);
                pLogFile.Dispose();
            }

        }

        public void WriteErrLog(string[] pMsg)
        {
            string vFlieName = "Err_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
            string vHeaderText = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "> ";

            CreateFolderLog();

            for (int i = 0; i < pMsg.Length; i++)
            {
                pMsg[i] = pMsg[i].PadLeft(pMsg[i].Length + vHeaderText.Length, ' ');    //padding with space character
                //switch(i)
                //{
                //    case 0:
                //        msg[i] = HeaderText + msg[i];
                //        break;
                //    default:
                //        msg[i] = msg[i].PadLeft(msg[i].Length + HeaderText.Length , ' ');    //padding with space character
                //        break;
                //}
            }
            using (System.IO.StreamWriter pLogFile = new StreamWriter(dirLog + vFlieName, true))
            {
                for (int i = 0; i < pMsg.Length; i++)
                {
                    pLogFile.WriteLine(pMsg[i]);
                }
                pLogFile.Dispose();
            }
        }

        private void ScanDeleteLog(string pPathLog, int pNumDay)
        {
            DateTime DateLastModified;
            DateTime DateDelete;
            try
            {
                DateDelete = DateTime.Now.AddDays(-1 * pNumDay);
                string[] fileEntries = Directory.GetFiles(pPathLog);
                foreach (string FileName in fileEntries)
                {
                    DateLastModified = File.GetCreationTime(FileName);
                    if (DateLastModified < DateDelete)
                    {
                        File.Delete(FileName);
                    }
                }
            }
            catch (Exception exp)
            { }
        }
    }
}
