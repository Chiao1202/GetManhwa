using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using HtmlAgilityPack;
using System.Net;
using System.IO;
using System.Drawing;
using System.Threading;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Collections;
using mshtml;
using ChromeBrowser;
using Microsoft.Win32;
using System.Security;
using System.Collections.Concurrent;
using System.Web;
using System.Diagnostics;
using System.Threading.Tasks;
namespace GetManhwa
{

    public partial class Form1 : Form
    {
        static int sleepSecond = 3;//載入圖時需要延遲
        public enum BrowserEmulationVersion
        {
            Default = 0,
            Version7 = 7000,
            Version8 = 8000,
            Version8Standards = 8888,
            Version9 = 9000,
            Version9Standards = 9999,
            Version10 = 10000,
            Version10Standards = 10001,
            Version11 = 11000,
            Version11Edge = 11001
        }
        public static class WBEmulator
        {
            private const string InternetExplorerRootKey = @"Software\Microsoft\Internet Explorer";

            public static int GetInternetExplorerMajorVersion()
            {
                int result;

                result = 0;

                try
                {
                    RegistryKey key;

                    key = Registry.LocalMachine.OpenSubKey(InternetExplorerRootKey);

                    if (key != null)
                    {
                        object value;

                        value = key.GetValue("svcVersion", null) ?? key.GetValue("Version", null);

                        if (value != null)
                        {
                            string version;
                            int separator;

                            version = value.ToString();
                            separator = version.IndexOf('.');
                            if (separator != -1)
                            {
                                int.TryParse(version.Substring(0, separator), out result);
                            }
                        }
                    }
                }
                catch (SecurityException)
                {
                    // The user does not have the permissions required to read from the registry key.
                }
                catch (UnauthorizedAccessException)
                {
                    // The user does not have the necessary registry rights.
                }

                return result;
            }
            private const string BrowserEmulationKey = InternetExplorerRootKey + @"\Main\FeatureControl\FEATURE_BROWSER_EMULATION";

            public static BrowserEmulationVersion GetBrowserEmulationVersion()
            {
                BrowserEmulationVersion result;

                result = BrowserEmulationVersion.Default;

                try
                {
                    RegistryKey key;

                    key = Registry.CurrentUser.OpenSubKey(BrowserEmulationKey, true);
                    if (key != null)
                    {
                        string programName;
                        object value;

                        programName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);
                        value = key.GetValue(programName, null);

                        if (value != null)
                        {
                            result = (BrowserEmulationVersion)Convert.ToInt32(value);
                        }
                    }
                }
                catch (SecurityException)
                {
                    // The user does not have the permissions required to read from the registry key.
                }
                catch (UnauthorizedAccessException)
                {
                    // The user does not have the necessary registry rights.
                }

                return result;
            }
            public static bool SetBrowserEmulationVersion(BrowserEmulationVersion browserEmulationVersion)
            {
                bool result;

                result = false;

                try
                {
                    RegistryKey key;

                    key = Registry.CurrentUser.OpenSubKey(BrowserEmulationKey, true);

                    if (key != null)
                    {
                        string programName;

                        programName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);

                        if (browserEmulationVersion != BrowserEmulationVersion.Default)
                        {
                            // if it's a valid value, update or create the value
                            key.SetValue(programName, (int)browserEmulationVersion, RegistryValueKind.DWord);
                        }
                        else
                        {
                            // otherwise, remove the existing value
                            key.DeleteValue(programName, false);
                        }

                        result = true;
                    }
                }
                catch (SecurityException)
                {
                    // The user does not have the permissions required to read from the registry key.
                }
                catch (UnauthorizedAccessException)
                {
                    // The user does not have the necessary registry rights.
                }

                return result;
            }

            public static bool SetBrowserEmulationVersion()
            {
                int ieVersion;
                BrowserEmulationVersion emulationCode;

                ieVersion = GetInternetExplorerMajorVersion();

                if (ieVersion >= 11)
                {
                    emulationCode = BrowserEmulationVersion.Version11;
                }
                else
                {
                    switch (ieVersion)
                    {
                        case 10:
                            emulationCode = BrowserEmulationVersion.Version10;
                            break;
                        case 9:
                            emulationCode = BrowserEmulationVersion.Version9;
                            break;
                        case 8:
                            emulationCode = BrowserEmulationVersion.Version8;
                            break;
                        default:
                            emulationCode = BrowserEmulationVersion.Version7;
                            break;
                    }
                }

                return SetBrowserEmulationVersion(emulationCode);
            }
            public static bool IsBrowserEmulationSet()
            {
                return GetBrowserEmulationVersion() != BrowserEmulationVersion.Default;
            }
        }
        public Form1()
        {

            InitializeComponent();
            //程序啟動時運行
            if (!WBEmulator.IsBrowserEmulationSet())
            {
                WBEmulator.SetBrowserEmulationVersion();
            }
            DisableClickSounds();//導頁靜音

        }
        bool IsRecord = true;//是否做記錄

        string OphenFileName = "OphenList.csv";
        string RecordFileName = "RecordList.csv";
        private void Form1_Load(object sender, EventArgs e)
        {
            //※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※
            //HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://i.hamreus.com/ps3/y/yiquanchaoren/%E7%AC%AC179%E8%AF%9D/0TIbMu.png.webp?e=1605310131&m=vzvYBcVxGKB-CsJzX5-XDQ");
            //request.Headers.Clear();
            //request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36";
            //request.Referer = "https://www.manhuagui.com/";
            //HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            //Stream stream = response.GetResponseStream();
            //byte[] Result = ReadStream(stream, 32765);
            //stream.Close();
            //FileStream myFile = File.Open(@"D:\可刪\AAA.png", FileMode.Create);
            //myFile.Write(Result, 0, Result.Length);
            //myFile.Dispose();
            //※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※※



            //升級WebBrowser版本
            WBEmulator.SetBrowserEmulationVersion();


            DataTable dt = new DataTable();


            DataTable dtTemp = dt.Clone();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                DateTime tstartTime;
                if (DateTime.TryParse(dr["startTime"].ToString().Trim(), out tstartTime))
                {
                    dr["endTime"] = tstartTime.AddHours(3).ToString("yyyy-MM-dd HH:mm:ss");
                    DataRow drNew = dtTemp.NewRow();
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (col.ColumnName.Trim().ToUpper() == "STARTTIME")
                            drNew[col.ColumnName] = tstartTime.AddHours(-3).ToString("yyyy-MM-dd HH:mm:ss");

                        else if (col.ColumnName.Trim().ToUpper() == "ENDTIME")
                            drNew[col.ColumnName] = tstartTime.ToString("yyyy-MM-dd HH:mm:ss");
                        else
                            drNew[col.ColumnName] = dr[col.ColumnName];
                    }
                    dtTemp.Rows.Add(drNew);
                }

            }
            dt.Merge(dtTemp);
            var v = from t in dt.AsEnumerable()
                    orderby t["locationName"].ToString().Trim(), t["startTime"].ToString().Trim()
                    select t;
            if (v != null && v.Count() > 0)
                dt = v.CopyToDataTable();

            var tDll = this.GetType().Assembly.GetName();
            string Version = tDll.Version.ToString();

            //自動更改版本號 愛看漫抓抓 VX.X
            this.Text = string.Format("愛看漫抓抓 V" + Version.Replace(".0", ""));

            #region 版本使用期限
            GetSettingItem g = new GetSettingItem();

            //時間驗證
            //加密範例
            //string KeyResult_ = AesEncrypt("2019-12-31|JOE", "GetManhwa");
            //string KeyResult2 = AesEncrypt("2020-01-31|NORMAL", "GetManhwa");
            bool Pass = true;
            DateTime NTPTime;
            if (GetNtpTime(8, out NTPTime))

                if (NTPTime >= Convert.ToDateTime("2018/01/30"))
                {
                    string ReleaseKey = g.GetSetValue("ReleaseKey");//ReleaseKey
                    if (ReleaseKey.Trim() == "")
                        Pass = false;
                    else
                    {
                        string KeyResult = "";
                        try
                        {
                            KeyResult = AesDecrypt(ReleaseKey, "GetManhwa");
                        }
                        catch (Exception ex)
                        {
                            Pass = false;

                        }
                        //2020-01-31|NORMAL
                        string[] strs = KeyResult.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                        if (strs.Length != 2)
                            Pass = false;
                        else
                        {
                            DateTime ReleaseTime;
                            if (!DateTime.TryParse(strs[0], out ReleaseTime))
                                Pass = false;
                            else
                            {
                                if (ReleaseTime <= NTPTime)
                                    Pass = false;
                            }
                        }
                    }

                }



            string ReleaseKey_New = AesEncrypt("2021-01-31|NORMAL", "GetManhwa");

            if (!Pass)
            {
                MessageBox.Show("感謝您的愛用!\r\n軟體為了持續更新，目前版本已到使用期限!\r\n請下載新版\r\n或請聯絡yusooyoun2@gmail.com;ㄚ僑\r\n", "使用說明", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.Close();
                return;
            }

            #endregion


            GetDefaultSavePath();//取得預設儲存路徑設定



            MessageBox.Show("感謝您的愛用!\r\n此軟體為程式軟體開發交流使用!\r\n請勿隨意散佈下載內容，以免觸法，謝謝。\r\n此軟體為C#語言開發，若對程式碼內容有興趣者請聯絡yusooyoun2@gmail.com;ㄚ僑\r\n", "使用說明", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            wbSearch.Navigated += wbSearch_Navigated;
            wbSearch.DocumentCompleted += wbSearch_DocumentCompleted;
            wbSearch.ScriptErrorsSuppressed = true;
            //wbSearch.Navigate("http://tw.manhuagui.com/");
            btnOpenClosePanel.PerformClick();




            //從CSV取得最愛項目
            CSVToListView(ref lvOphen, ref imlItem, OphenFileName);

            //從CSV取得記錄項目
            CSVToListView(ref lvRecord, ref imlRecord, RecordFileName);

            //縮起介紹
            scpOphen.Panel2Collapsed = true;

            //縮起
            scRecord.Panel2Collapsed = true;

            txtURL.Focus();
            txtURL.SelectAll();

            //下載延遲
            cmbSleep_S.SelectedIndex = 2;
            cmbSleep_E.SelectedIndex = 3;
        }


        public byte[] ReadStream(Stream stream, int initialLength)
        {
            if (initialLength < 1)
            {
                initialLength = 32768;
            }
            byte[] buffer = new byte[initialLength];
            int read = 0;
            int chunk;
            while ((chunk = stream.Read(buffer, read, buffer.Length - read)) > 0)
            {
                read += chunk;
                if (read == buffer.Length)
                {
                    int nextByte = stream.ReadByte();
                    if (nextByte == -1)
                    {
                        return buffer;
                    }
                    byte[] newBuffer = new byte[buffer.Length * 2];
                    Array.Copy(buffer, newBuffer, buffer.Length);
                    newBuffer[read] = (byte)nextByte;
                    buffer = newBuffer;
                    read++;
                }
            }
            byte[] bytes = new byte[read];
            Array.Copy(buffer, bytes, read);
            return bytes;
        }

        private bool GetNtpTime(int iUTCAdd, out DateTime networkDateTime)
        {
            networkDateTime = DateTime.Now;
            try
            {

                //const string ntpServer = "time.nist.gov";
                const string ntpServer = "tick.stdtime.gov.tw";

                byte[] ntpData = new byte[48];

                //LeapIndicator = 0 (no warning), VersionNum = 3 (IPv4 only), Mode = 3 (Client Mode)
                ntpData[0] = 0x1B;

                IPAddress[] addresses = Dns.GetHostEntry(ntpServer).AddressList;
                IPEndPoint ipEndPoint = new IPEndPoint(addresses[0], 123);
                Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);

                socket.Connect(ipEndPoint);
                socket.Send(ntpData);
                socket.Receive(ntpData);
                socket.Close();


                const byte serverReplyTime = 40;
                //Get the seconds part
                ulong intPart = BitConverter.ToUInt32(ntpData, serverReplyTime);

                //Get the seconds fraction
                ulong fractPart = BitConverter.ToUInt32(ntpData, serverReplyTime + 4);

                //Convert From big-endian to little-endian
                intPart = SwapEndianness(intPart);
                fractPart = SwapEndianness(fractPart);

                var milliseconds = (intPart * 1000) + ((fractPart * 1000) / 0x100000000L);

                //UTC time + 8 
                networkDateTime = (new DateTime(1900, 1, 1))
                    .AddMilliseconds((long)milliseconds).AddHours(iUTCAdd);

                //UnityEngine.Debug.Log(networkDateTime);

                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        static uint SwapEndianness(ulong x)
        {
            return (uint)(((x & 0x000000ff) << 24) +
                          ((x & 0x0000ff00) << 8) +
                          ((x & 0x00ff0000) >> 8) +
                          ((x & 0xff000000) >> 24));
        }
        private void GetDefaultSavePath()
        {
            try
            {
                GetSettingItem g = new GetSettingItem();



                string IsUseDefaultSavePath = g.GetSetValue("IsUseDefaultSavePath");//是否使用預設儲存路徑
                string DefaultSavePath = g.GetSetValue("DefaultSavePath");//預設儲存路徑

                if (IsUseDefaultSavePath == "1")//使用預設儲存路徑
                {
                    chkIsUseDefaultSavePath.Checked = true;
                    txtDirPath.Enabled = true;
                }
                else
                {
                    chkIsUseDefaultSavePath.Checked = false;
                    txtDirPath.Enabled = false;
                }

                if (DefaultSavePath.Trim() != "")//預設儲存路徑
                    txtDirPath.Text = DefaultSavePath;



                //下載記錄
                string IsRecord = g.GetSetValue("IsRecord");//預設儲存路徑
                if (IsRecord.Trim() != "")
                {
                    if (IsRecord.Trim().ToUpper() == "TRUE")
                        chkRecord.Checked = true;
                    else
                        chkRecord.Checked = false;
                }


                string ScanWidth_ = g.GetSetValue("ScanWidth");//判斷的螢幕寬
                string ScanHeight_ = g.GetSetValue("ScanHeight");//判斷的螢幕高
                if (!int.TryParse(ScanWidth_, out ScanWidth))
                    ScanWidth = 2500;
                if (!int.TryParse(ScanHeight_, out ScanHeight))
                    ScanHeight = 2500;



            }
            catch (Exception ex)
            {

            }
        }

        public static string Decrypt(string Text, string strKey)
        {
            return Decrypt(Text, strKey, Encoding.Default);
        }
        /// <summary> 
        /// 解密数据 
        /// </summary> 
        /// <param name="Text"></param> 
        /// <param name="strKey"></param> 
        /// <returns></returns> 
        public static string Decrypt(string Text, string strKey, Encoding encoding)
        {
            DESCryptoServiceProvider des = new DESCryptoServiceProvider();
            int len;
            len = Text.Length / 2;
            byte[] inputByteArray = new byte[len];
            int x, i;
            for (x = 0; x < len; x++)
            {
                i = Convert.ToInt32(Text.Substring(x * 2, 2), 16);
                inputByteArray[x] = (byte)i;
            }
            des.Key = ASCIIEncoding.ASCII.GetBytes(System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(strKey, "md5").Substring(0, 8));
            des.IV = ASCIIEncoding.ASCII.GetBytes(System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(strKey, "md5").Substring(0, 8));
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();

            ms.Close();
            cs.Close();
            return encoding.GetString(ms.ToArray());
        }



        string strURL;
        string P1;
        int P_S;
        int P_E;
        int Index;
        static int ScanWidth, ScanHeight;//截螢幕畫面時原始的寬高(影像越大，寬高要越大，太大也會影響效能)
        string DirPath = "";
        private void btnGet_Click(object sender, EventArgs e)
        {
            #region 取得存檔資料夾路徑
            string DirPath = "";
            if (chkIsUseDefaultSavePath.Checked)//使用預設儲存路徑
            {
                if (txtDirPath.Text.Trim() == "")//無設定時
                {
                    MessageBox.Show("預設存檔路徑錯誤!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                    txtDirPath.Focus();
                    txtDirPath.SelectAll();
                    return;
                }
                else
                {
                    if (!Directory.Exists(txtDirPath.Text.Trim()))
                    {
                        try
                        {
                            Directory.CreateDirectory(txtDirPath.Text.Trim());
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("預設存檔路徑錯誤!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                            return;
                        }
                    }
                    else
                        DirPath = txtDirPath.Text.Trim();
                }
            }
            else
            {
                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                    DirPath = folderBrowserDialog1.SelectedPath;
                else
                    return;
            }
            if (DirPath.LastIndexOf("\\") != DirPath.Length - 1)
                DirPath += "\\";
            #endregion

            #region 內頁重覆的處理方式
            int SaveProcessType = 0;
            if (rdbContinue.Checked)
                SaveProcessType = 0;
            if (rdbCover.Checked)
                SaveProcessType = 1;
            if (rdbRename.Checked)
                SaveProcessType = 2;
            #endregion

            #region 起始頁 結束頁
            int P_S = 1;
            if (txtP_S.Text.Trim() == "" || int.TryParse(txtP_S.Text.Trim(), out P_S))
            {
                MessageBox.Show("起始頁錯誤!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                txtP_S.Focus();
                txtP_S.SelectAll();
                return;
            }
            int P_E = 0;
            if (!int.TryParse(txtP_E.Text.Trim(), out P_E))
                P_E = 0;
            #endregion

            P1 = "#p=";

            string URL = txtURL.Text.Trim();
            string Contents = "";
            try
            {
                WebClient wc = new WebClient();
                byte[] bResult = wc.DownloadData(strURL);
                Contents = Encoding.UTF8.GetString(bResult);
            }
            catch (Exception ex)
            {
                MessageBox.Show("錯誤", ex.Message);
                return;
            }
            //(未完成)
            //依Contents取得漫畫名、卷名及總頁數
            int PageCount = 0;

            #region 取得漫畫名稱
            string MainName = "";
            if (txtMainName.Text.Trim() == "")
            {
                MessageBox.Show("漫畫名稱錯誤!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                txtMainName.Focus();
                txtMainName.SelectAll();
                return;
            }
            else
                MainName = txtMainName.Text.Trim();

            if (MainName.LastIndexOf("\\") != MainName.Length - 1)
                MainName += "\\";
            #endregion



            int Index = 1;
            if (listWorkOption != null && listWorkOption.Count > 0)
                Index = listWorkOption.Max(t => t.NO) + 1;


            AddWork(Index, URL, DirPath, MainName, Name, PageCount, ScanWidth, ScanHeight, SaveProcessType);
            ShowWorkList();//顯示項目
            //CheckWorkAndStartWork();//檢查是否有未執行的工作並執行工作ShowWorkList(取完所有清單後)

        }


        public void loading()
        {
            DateTime dateNow = DateTime.Now;
            try
            {
                while (!(wb.ReadyState == WebBrowserReadyState.Complete) && (DateTime.Now - dateNow).TotalSeconds <= 10)
                    Application.DoEvents();
            }
            catch (Exception ex)
            {

            }
        }

        private string GetContentType(string FE)
        {
            #region
            switch (FE.Trim().ToUpper())
            {
                case "001": return "application/x-001";
                case "323": return "text/h323";
                case "907": return "drawing/907";
                case "ACP": return "audio/x-mei-aac";
                case "AIF": return "audio/aiff";
                case "AIFF": return "audio/aiff";
                case "ASA": return "text/asa";
                case "ASP": return "text/asp";
                case "AU": return "audio/basic";
                case "AWF": return "application/vnd.adobe.workflow";
                case "BMP": return "application/x-bmp";
                case "C4T": return "application/x-c4t";
                case "CAL": return "application/x-cals";
                case "CDF": return "application/x-netcdf";
                case "CEL": return "application/x-cel";
                case "CG4": return "application/x-g4";
                case "CIT": return "application/x-cit";
                case "CML": return "text/xml";
                case "CMX": return "application/x-cmx";
                case "CRL": return "application/pkix-crl";
                case "CSI": return "application/x-csi";
                case "CUT": return "application/x-cut";
                case "DBM": return "application/x-dbm";
                case "DCD": return "text/xml";
                case "DER": return "application/x-x509-ca-cert";
                case "DIB": return "application/x-dib";
                case "DOC": return "application/msword";
                case "DRW": return "application/x-drw";
                //case "dwf": return "Model/vnd.dwf";
                case "DWG": return "application/x-dwg";
                case "DXF": return "application/x-dxf";
                case "EMF": return "application/x-emf";
                case "ENT": return "text/xml";
                //case "eps": return "application/x-ps";
                case "ETD": return "application/x-ebx";
                case "FAX": return "image/fax";
                case "FIF": return "application/fractals";
                case "FRM": return "application/x-frm";
                case "GBR": return "application/x-gbr";
                case "GIF": return "image/gif";
                case "GP4": return "application/x-gp4";
                case "HMR": return "application/x-hmr";
                case "HPL": return "application/x-hpl";
                case "HRF": return "application/x-hrf";
                case "HTC": return "text/x-component";
                case "HTML": return "text/html";
                case "HTX": return "text/html";
                case "ICO": return "image/x-icon";
                case "IFF": return "application/x-iff";
                case "IGS": return "application/x-igs";
                case "IMG": return "application/x-img";
                case "ISP": return "application/x-internet-signup";
                case "JAVA": return "java/*";
                case "JPE": return "image/jpeg";
                case "JPEG": return "image/jpeg";
                //case "jpg": return "application/x-jpg";
                case "JSP": return "text/html";
                case "LAR": return "application/x-laplayer-reg";
                case "LAVS": return "audio/x-liquid-secure";
                case "LMSFF": return "audio/x-la-lms";
                case "LTR": return "application/x-ltr";
                case "M2V": return "video/x-mpeg";
                case "M4E": return "video/mpeg4";
                case "MAN": return "application/x-troff-man";
                case "MDB": return "application/msaccess";
                case "MFP": return "application/x-shockwave-flash";
                case "MHTML": return "message/rfc822";
                case "MID": return "audio/mid";
                case "MIL": return "application/x-mil";
                case "MND": return "audio/x-musicnet-download";
                case "MOCHA": return "application/x-javascript";
                case "MP1": return "audio/mp1";
                case "MP2V": return "video/mpeg";
                case "MP4": return "video/mpeg4";
                case "MPD": return "application/vnd.ms-project";
                case "MPEG": return "video/mpg";
                case "MPGA": return "audio/rn-mpeg";
                case "MPS": return "video/x-mpeg";
                case "MPV": return "video/mpg";
                case "MPW": return "application/vnd.ms-project";
                case "MTX": return "text/xml";
                case "NET": return "image/pnetvue";
                case "NWS": return "message/rfc822";
                case "OUT": return "application/x-out";
                case "P12": return "application/x-pkcs12";
                case "P7C": return "application/pkcs7-mime";
                case "P7R": return "application/x-pkcs7-certreqresp";
                case "PC5": return "application/x-pc5";
                case "PCL": return "application/x-pcl";
                case "PDF": return "application/pdf";
                case "PDX": return "application/vnd.adobe.pdx";
                case "PGL": return "application/x-pgl";
                case "PKO": return "application/vnd.ms-pki.pko";
                case "PLG": return "text/html";
                case "PLT": return "application/x-plt";
                //case "png": return "application/x-png";
                case "PPA": return "application/vnd.ms-powerpoint";
                case "PPS": return "application/vnd.ms-powerpoint";
                //case "ppt": return "application/x-ppt";
                case "PRF": return "application/pics-rules";
                case "PRT": return "application/x-prt";
                case "PS": return "application/postscript";
                case "PWZ": return "application/vnd.ms-powerpoint";
                case "RA": return "audio/vnd.rn-realaudio";
                case "RAS": return "application/x-ras";
                case "RDF": return "text/xml";
                case "RED": return "application/x-red";
                case "RJS": return "application/vnd.rn-realsystem-rjs";
                case "RLC": return "application/x-rlc";
                case "RM": return "application/vnd.rn-realmedia";
                case "RMI": return "audio/mid";
                case "RMM": return "audio/x-pn-realaudio";
                case "RMS": return "application/vnd.rn-realmedia-secure";
                case "RMX": return "application/vnd.rn-realsystem-rmx";
                case "RP": return "image/vnd.rn-realpix";
                case "RSML": return "application/vnd.rn-rsml";
                case "RTF": return "application/msword";
                case "RV": return "video/vnd.rn-realvideo";
                case "SAT": return "application/x-sat";
                case "SDW": return "application/x-sdw";
                case "SLB": return "application/x-slb";
                case "SLK": return "drawing/x-slk";
                case "SMIL": return "application/smil";
                case "SND": return "audio/basic";
                case "SOR": return "text/plain";
                case "SPL": return "application/futuresplash";
                case "SSM": return "application/streamingmedia";
                case "STL": return "application/vnd.ms-pki.stl";
                case "STY": return "application/x-sty";
                case "SWF": return "application/x-shockwave-flash";
                case "TG4": return "application/x-tg4";
                case "TIF": return "image/tiff";
                case "TIFF": return "image/tiff";
                case "TOP": return "drawing/x-top";
                case "TSD": return "text/xml";
                case "UIN": return "application/x-icq";
                case "VCF": return "text/x-vcard";
                case "VDX": return "application/vnd.visio";
                case "VPG": return "application/x-vpeg005";
                //case "vsd": return "application/x-vsd";
                case "VST": return "application/vnd.visio";
                case "VSW": return "application/vnd.visio";
                case "VTX": return "application/vnd.visio";
                case "WAV": return "audio/wav";
                case "WB1": return "application/x-wb1";
                case "WB3": return "application/x-wb3";
                case "WIZ": return "application/msword";
                case "WK4": return "application/x-wk4";
                case "WKS": return "application/x-wks";
                case "WMA": return "audio/x-ms-wma";
                case "WMF": return "application/x-wmf";
                case "WMV": return "video/x-ms-wmv";
                case "WMZ": return "application/x-ms-wmz";
                case "WPD": return "application/x-wpd";
                case "WPL": return "application/vnd.ms-wpl";
                case "WR1": return "application/x-wr1";
                case "WRK": return "application/x-wrk";
                case "WS2": return "application/x-ws";
                case "WSDL": return "text/xml";
                case "XDP": return "application/vnd.adobe.xdp";
                case "XFD": return "application/vnd.adobe.xfd";
                case "XHTML": return "text/html";
                //case "xls": return "application/x-xls";
                case "XML": return "text/xml";
                case "XQ": return "text/xml";
                case "XQUERY": return "	text/xml";
                case "XSL": return "text/xml";
                case "XWD": return "application/x-xwd";
                case "SIS": return "application/vnd.symbian.install";
                case "X_T": return "application/x-x_t";
                case "APK": return "application/vnd.android.package-archive";
                case "301": return "application/x-301";
                case "906": return "application/x-906";
                case "A11": return "application/x-a11";
                case "AI": return "application/postscript";
                case "AIFC": return "audio/aiff";
                case "ANV": return "application/x-anv";
                case "ASF": return "video/x-ms-asf";
                case "ASX": return "video/x-ms-asf";
                case "AVI": return "video/avi";
                case "BIZ": return "text/xml";
                case "BOT": return "application/x-bot";
                case "C90": return "application/x-c90";
                case "CAT": return "application/vnd.ms-pki.seccat";
                case "CDR": return "application/x-cdr";
                case "CER": return "application/x-x509-ca-cert";
                case "CGM": return "application/x-cgm";
                case "CLASS": return "java/*";
                case "CMP": return "application/x-cmp";
                case "COT": return "application/x-cot";
                case "CRT": return "application/x-x509-ca-cert";
                case "CSS": return "text/css";
                case "DBF": return "application/x-dbf";
                case "DBX": return "application/x-dbx";
                case "DCX": return "application/x-dcx";
                case "DGN": return "application/x-dgn";
                case "DLL": return "application/x-msdownload";
                case "DOT": return "application/msword";
                case "DTD": return "text/xml";
                case "DWF": return "application/x-dwf";
                case "DXB": return "application/x-dxb";
                case "EDN": return "application/vnd.adobe.edn";
                case "EML": return "message/rfc822";
                case "EPI": return "application/x-epi";
                case "EPS": return "application/postscript";
                case "EXE": return "application/x-msdownload";
                case "FDF": return "application/vnd.fdf";
                case "FO": return "text/xml";
                case "G4": return "application/x-g4";
                case "GL2": return "application/x-gl2";
                case "HGL": return "application/x-hgl";
                case "HPG": return "application/x-hpgl";
                case "HQX": return "application/mac-binhex40";
                case "HTA": return "application/hta";
                case "HTM": return "text/html";
                case "HTT": return "text/webviewhtml";
                case "ICB": return "application/x-icb";
                //case "ico": return "application/x-ico";
                case "IG4": return "application/x-g4";
                case "III": return "application/x-iphone";
                case "INS": return "application/x-internet-signup";
                case "IVF": return "video/x-ivf";
                case "JFIF": return "image/jpeg";
                //case "jpe": return "application/x-jpe";
                case "JPG": return "image/jpeg";
                case "JS": return "application/x-javascript";
                case "LA1": return "audio/x-liquid-file";
                case "LATEX": return "application/x-latex";
                case "LBM": return "application/x-lbm";
                case "LS": return "application/x-javascript";
                case "M1V": return "video/x-mpeg";
                case "M3U": return "audio/mpegurl";
                case "MAC": return "application/x-mac";
                case "MATH": return "text/xml";
                //case "mdb": return "application/x-mdb";
                case "MHT": return "message/rfc822";
                case "MI": return "application/x-mi";
                case "MIDI": return "audio/mid";
                case "MML": return "text/xml";
                case "MNS": return "audio/x-musicnet-stream";
                case "MOVIE": return "video/x-sgi-movie";
                case "MP2": return "audio/mp2";
                case "MP3": return "audio/mp3";
                case "MPA": return "video/x-mpg";
                case "MPE": return "video/x-mpeg";
                case "MPG": return "video/mpg";
                case "MPP": return "application/vnd.ms-project";
                case "MPT": return "application/vnd.ms-project";
                case "MPV2": return "video/mpeg";
                case "MPX": return "application/vnd.ms-project";
                case "MXP": return "application/x-mmxp";
                case "NRF": return "application/x-nrf";
                case "ODC": return "text/x-ms-odc";
                case "P10": return "application/pkcs10";
                case "P7B": return "application/x-pkcs7-certificates";
                case "P7M": return "application/pkcs7-mime";
                case "P7S": return "application/pkcs7-signature";
                case "PCI": return "application/x-pci";
                case "PCX": return "application/x-pcx";
                //case "pdf": return "application/pdf";
                case "PFX": return "application/x-pkcs12";
                case "PIC": return "application/x-pic";
                case "PL": return "application/x-perl";
                case "PLS": return "audio/scpls";
                case "PNG": return "image/png";
                case "POT": return "application/vnd.ms-powerpoint";
                case "PPM": return "application/x-ppm";
                case "PPT": return "application/vnd.ms-powerpoint";
                case "PR": return "application/x-pr";
                case "PRN": return "application/x-prn";
                //case "ps": return "application/x-ps";
                case "PTN": return "application/x-ptn";
                case "R3T": return "text/vnd.rn-realtext3d";
                case "RAM": return "audio/x-pn-realaudio";
                case "RAT": return "application/rat-file";
                case "REC": return "application/vnd.rn-recording";
                case "RGB": return "application/x-rgb";
                case "RJT": return "application/vnd.rn-realsystem-rjt";
                case "RLE": return "application/x-rle";
                case "RMF": return "application/vnd.adobe.rmf";
                case "RMJ": return "application/vnd.rn-realsystem-rmj";
                case "RMP": return "application/vnd.rn-rn_music_package";
                case "RMVB": return "application/vnd.rn-realmedia-vbr";
                case "RNX": return "application/vnd.rn-realplayer";
                case "RPM": return "audio/x-pn-realaudio-plugin";
                case "RT": return "text/vnd.rn-realtext";
                //case "rtf": return "application/x-rtf";
                case "SAM": return "application/x-sam";
                case "SDP": return "application/sdp";
                case "SIT": return "application/x-stuffit";
                case "SLD": return "application/x-sld";
                case "SMI": return "application/smil";
                case "SMK": return "application/x-smk";
                case "SOL": return "text/plain";
                case "SPC": return "application/x-pkcs7-certificates";
                case "SPP": return "text/xml";
                case "SST": return "application/vnd.ms-pki.certstore";
                case "STM": return "text/html";
                case "SVG": return "text/xml";
                case "TDF": return "application/x-tdf";
                case "TGA": return "application/x-tga";
                //case "tif": return "application/x-tif";
                case "TLD": return "text/xml";
                case "TORRENT": return "application/x-bittorrent";
                case "TXT": return "text/plain";
                case "ULS": return "text/iuls";
                case "VDA": return "application/x-vda";
                case "VML": return "text/xml";
                case "VSD": return "application/vnd.visio";
                case "VSS": return "application/vnd.visio";
                //case "vst": return "application/x-vst";
                case "VSX": return "application/vnd.visio";
                case "VXML": return "text/xml";
                case "WAX": return "audio/x-ms-wax";
                case "WB2": return "application/x-wb2";
                case "WBMP": return "image/vnd.wap.wbmp";
                case "WK3": return "application/x-wk3";
                case "WKQ": return "application/x-wkq";
                case "WM": return "video/x-ms-wm";
                case "WMD": return "application/x-ms-wmd";
                case "WML": return "text/vnd.wap.wml";
                case "WMX": return "video/x-ms-wmx";
                case "WP6": return "application/x-wp6";
                case "WPG": return "application/x-wpg";
                case "WQ1": return "application/x-wq1";
                case "WRI": return "application/x-wri";
                case "WS": return "application/x-ws";
                case "WSC": return "text/scriptlet";
                case "WVX": return "video/x-ms-wvx";
                case "XDR": return "text/xml";
                case "XFDF": return "application/vnd.adobe.xfdf";
                case "XLS": return "application/vnd.ms-excel";
                case "XLW": return "application/x-xlw";
                case "XPL": return "audio/scpls";
                case "XQL": return "text/xml";
                case "XSD": return "text/xml";
                case "XSLT": return "text/xml";
                case "X_B": return "application/x-x_b";
                case "SISX": return "application/vnd.symbian.install";
                case "IPA": return "application/vnd.iphone";
                case "XAP": return "application/x-silverlight-app";
                case "ZIP": return "application/x-zip-compressed";
                case "JSON": return "application/json";
                default:
                    return "";
            }

            #endregion

        }

        private void btnSetting_Click(object sender, EventArgs e)//取得所有集數資料
        {
            if (txtURL.Text.Trim() == "")
            {
                MessageBox.Show("請輸入連結!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                return;
            }


            string strURL_O = txtURL.Text.Trim();//http://tw.ikanman.com/comic/20555/  http://tw.manhuagui.com/comic/20275/
            strURL_O = strURL_O.Replace("https:", "http:");
            int NUM;
            if (int.TryParse(strURL_O, out NUM))
            {
                strURL_O = string.Format("http://tw.ikanman.com/comic/{0}", NUM.ToString());
                txtURL.Text = strURL_O;
            }

            string strNUM = "";
            string strSUB_NUM = "";
            string P = "";
            GetNumURL(strURL_O, out strURL, out strNUM, out strSUB_NUM, out P);
            txtURL.Text = strURL;

            DataTable dt = new DataTable();
            string Titlt;
            dt = GetAllBookDataTable(strURL, "", out Titlt);
            txtMainName.Text = Titlt;



            if (strSUB_NUM.Trim() != "")
            {
                var v = from t in dt.AsEnumerable()
                        where t["URL"].ToString().Trim().ToUpper().IndexOf("/" + strSUB_NUM.Trim() + ".HTML") > -1
                        select t;
                if (v != null && v.Count() > 0)
                    dt = v.CopyToDataTable();
            }
            gvList.DataSource = dt;
            if (gvList.Columns.Count > 4)
            {
                for (int i = 4; i < gvList.Columns.Count; i++)
                {
                    gvList.Columns[i].Visible = false;
                }
            }
            dtTemp = dt;
        }

        private void checkAdultClick(ref WebBrowser WB)
        {

            HtmlElement head = WB.Document.GetElementsByTagName("head")[0];
            HtmlElement scriptEl = WB.Document.CreateElement("script");
            //scriptEl.InnerHtml = "function Adult() { document.getElementById('checkAdult').click(); }"; ;
            IHTMLScriptElement element = (IHTMLScriptElement)scriptEl.DomElement;
            element.text = "function Adult() { document.getElementById('checkAdult').click(); }";
            head.AppendChild(scriptEl);
            WB.Document.InvokeScript("Adult");
            int tryCount = 0;
            loading(WB, ref tryCount);//等待完成載入
            //WB.Document.InvokeScript("document.getElementById('checkAdult').click();");

        }

        //處理連結
        private void GetNumURL(string strURL, out string strURL_out, out string strNUM, out string strSUB_NUM, out string p)
        {
            strURL_out = "";
            strSUB_NUM = "";
            p = "";
            strNUM = "";

            List<string> listKey = new List<string>();
            listKey.Add("http://tw.ikanman.com/comic/");
            listKey.Add("http://tw.manhuagui.com/comic/");
            listKey.Add("http://www.ikanman.com/comic/");
            listKey.Add("http://www.manhuagui.com/comic/");
            listKey.Add("https://tw.ikanman.com/comic/");
            listKey.Add("https://tw.manhuagui.com/comic/");
            listKey.Add("https://www.ikanman.com/comic/");
            listKey.Add("https://www.manhuagui.com/comic/");
            listKey.Add("/comic/");
            for (int i = 0; i < listKey.Count; i++)
            {

                int NUM;
                string Key_S = listKey[i].Trim();
                string Key_E = "/";

                int Index01 = strURL.IndexOf(Key_S);
                if (Index01 > -1)
                {
                    int URLLength = strURL.Length - Index01 - Key_S.Length;
                    int Index02 = strURL.LastIndexOf(Key_E);
                    if (Index02 == strURL.Length - 1)
                        URLLength -= 1;
                    strNUM = strURL.Substring(Index01 + Key_S.Length, URLLength);

                    //http://tw.manhuagui.com/comic/17023/250853.html#p=2
                    int indexHtml = strNUM.ToUpper().Trim().IndexOf(".HTML");
                    if (indexHtml > -1)
                    {
                        p = strNUM.Substring(indexHtml).ToUpper().Replace(".HTML", "").Replace("#P=", "");
                        strNUM = strNUM.Substring(0, indexHtml);
                    }
                    strNUM = strNUM.ToUpper().Replace(".HTML", "");
                    if (strNUM.IndexOf("/") > -1)
                    {
                        string[] strS = strNUM.Split('/');
                        strNUM = strS[0];
                        strSUB_NUM = strS[1];
                    }
                    if (int.TryParse(strNUM, out NUM))
                    {
                        strURL_out = string.Format("{0}{1}/", Key_S, NUM.ToString());
                        strNUM = NUM.ToString();
                    }
                    break;
                }
            }

            //strURL_out = "";
            //strSUB_NUM = "";
            //p = "";
            //int NUM;
            //strNUM = "";
            //string Key_S = "http://tw.ikanman.com/comic/";
            //string Key_E = "/";

            //int Index01 = strURL.IndexOf(Key_S);
            //if (Index01 > -1)
            //{
            //    int URLLength = strURL.Length - Index01 - Key_S.Length;
            //    int Index02 = strURL.LastIndexOf(Key_E);
            //    if (Index02 == strURL.Length - 1)
            //        URLLength -= 1;
            //    strNUM = strURL.Substring(Index01 + Key_S.Length, URLLength);

            //    //http://tw.manhuagui.com/comic/17023/250853.html#p=2
            //    int indexHtml = strNUM.ToUpper().Trim().IndexOf(".HTML");
            //    if (indexHtml > -1)
            //    {
            //        p = strNUM.Substring(indexHtml).ToUpper().Replace(".HTML", "").Replace("#P=", "");
            //        strNUM = strNUM.Substring(0, indexHtml);
            //    }
            //    strNUM = strNUM.ToUpper().Replace(".HTML", "");
            //    if (strNUM.IndexOf("/") > -1)
            //    {
            //        string[] strS = strNUM.Split('/');
            //        strNUM = strS[0];
            //        strSUB_NUM = strS[1];
            //    }
            //    if (int.TryParse(strNUM, out NUM))
            //    {
            //        strURL_out = string.Format("{0}{1}/", Key_S, NUM.ToString());
            //        strNUM = NUM.ToString();
            //    }
            //}
            //else
            //{
            //    Key_S = "http://tw.manhuagui.com/comic/";
            //    Key_E = "/";
            //    int Index03 = strURL.IndexOf(Key_S);
            //    if (Index03 > -1)
            //    {
            //        int URLLength = strURL.Length - Index03 - Key_S.Length;

            //        int Index02 = strURL.LastIndexOf(Key_E);
            //        if (Index02 == strURL.Length - 1)
            //            URLLength -= 1;
            //        strNUM = strURL.Substring(Index03 + Key_S.Length, URLLength);
            //        int indexHtml = strNUM.ToUpper().Trim().IndexOf(".HTML");
            //        if (indexHtml > -1)
            //        {
            //            p = strNUM.Substring(indexHtml).ToUpper().Replace(".HTML","").Replace("#P=", "");
            //            strNUM = strNUM.Substring(0, indexHtml);
            //        }
            //        strNUM = strNUM.ToUpper().Replace(".HTML", "");
            //        if (strNUM.IndexOf("/") > -1)
            //        {
            //            string[] strS = strNUM.Split('/');
            //            strNUM = strS[0];
            //            strSUB_NUM = strS[1];
            //        }
            //        if (int.TryParse(strNUM, out NUM))
            //        {
            //            strURL_out = string.Format("{0}{1}/", Key_S, NUM.ToString());
            //            strNUM = NUM.ToString();

            //        }
            //    }

            //}


        }

        private void getAllLink(string strURL, string Contents, ref DataTable dt)//遞迴取得所有集數的資料
        {
            if (Contents.IndexOf("<ul") < 0) return;//若未包含則跳過

            string Value01, Value02, NewContents;//所有單行本

            int Index1 = Contents.IndexOf("<ul style=\"display:block\">");
            int Index2 = Contents.IndexOf("<ul>");
            if (Index1 < Index2)
                GetValue(Contents, out NewContents, out Value01, "<ul style=\"display:block\">", "</ul>");//取得內部所有資料
            else
                GetValue(Contents, out NewContents, out Value01, "<ul>", "</ul>");//取得內部所有資料



            GetliInfo(strURL, Value01, ref dt);//依內部資料取得連結
            getAllLink(strURL, NewContents, ref dt);//再繼續挖資料
        }
        //取得該ul內所有資訊
        private void GetliInfo(string strURL_O, string ul, ref DataTable dt)
        {
            if (ul.IndexOf("<a href=\"") < 0 || ul.IndexOf("<span>") < 0 || ul.IndexOf("<i>") < 0) return;

            string URL;
            //取得連結
            string NewContents;//取得連結後的內容
            GetValue(ul, out NewContents, out URL, "<a href=\"", "\"");
            //取得卷名
            string Name;//卷名
            GetValue(NewContents, out Name, "<span>", "<i>");
            //取得頁數
            string PageCount;
            string NewContents2;//取得頁數後的內容
            GetValue(NewContents, out NewContents2, out PageCount, "<i>", "</i>");


            string[] OrgURLs = strURL_O.Split(new string[] { "/" }, StringSplitOptions.None);
            string[] URLs = URL.Split(new string[] { "/" }, StringSplitOptions.None);
            string CombinURL = "";
            for (int i = 0; i < URLs.Length; i++)
            {
                if (URLs[i].Trim() == "") continue;
                int IndexFlg = OrgURLs.ToList().FindIndex(U => U.Trim() == URLs[i].Trim() && U.Trim() != "");
                if (IndexFlg >= 0)
                {
                    CombinURL = string.Join("/", OrgURLs, 0, IndexFlg) + "/" + string.Join("/", URLs, i, URLs.Length - i);
                    break;
                }

            }
            string strNUM = "";
            string strSUB_NUM = "";
            string P = "";
            string strURL = "";
            GetNumURL(CombinURL, out strURL, out strNUM, out strSUB_NUM, out P);

            DataRow dr = dt.NewRow();
            dr["SEL"] = true;
            dr["Name"] = Name;
            dr["PageCount"] = PageCount.ToUpper().Trim().Replace("P", "");
            dr["URL"] = strURL + strSUB_NUM + ".html";//CombinURL;
            dr["SubNum"] = strSUB_NUM;
            dr["IsDownLoad"] = 0;
            dr["Status"] = "未下載";
            dt.Rows.Add(dr);

            if (NewContents2.IndexOf("<a href=\"") >= 0)
                GetliInfo(strURL_O, NewContents2, ref dt);

        }

        //取得該ul內所有資訊
        private void GetliInfo2(string strURL_O, string ul, ref DataTable dt)
        {
            if (ul.ToUpper().IndexOf("<A ") < 0 || ul.ToUpper().IndexOf("HREF") < 0 || ul.ToUpper().IndexOf("<SPAN>") < 0 || ul.ToUpper().IndexOf("<I>") < 0) return;
            string NewContents;//取得連結後的內容

            string Name;//卷名
            int Index01 = ul.IndexOf("title=\"");
            int Index02 = ul.IndexOf("title=");


            if (Index02 < Index01 || Index01 == -1)
                GetValue(ul, out NewContents, out Name, "title=", " ");
            else
                GetValue(ul, out NewContents, out Name, "title=\"", "\"");


            if (Name.IndexOf("\"") > -1)
            {

            }
            string URL;//取得連結
            GetValue(NewContents, out NewContents, out URL, "href=\"", "\"");
            string strNUM = "";
            string strSUB_NUM = "";
            string P = "";
            GetNumURL(URL, out strURL, out strNUM, out strSUB_NUM, out P);


            //取得卷名
            if (Name.Trim() == "")
                GetValue(NewContents, out Name, "<span>", "<i>");
            //取得頁數
            string PageCount;
            string NewContents2;//取得頁數後的內容
            GetValue(NewContents, out NewContents2, out PageCount, "<i>", "</i>");


            int intPageCount = 0;
            if (!int.TryParse(PageCount.ToUpper().Trim().Replace("P", ""), out intPageCount))
                intPageCount = 0;

            if (Name.Trim() != "" && PageCount.Trim() != "" && strSUB_NUM.Trim() != "")
            {
                if (strURL.IndexOf("http") < 0)//連結不完整時
                {
                    strURL = strURL_O;
                }

                DataRow dr = dt.NewRow();
                dr["SEL"] = true;
                dr["Name"] = Name;
                dr["PageCount"] = intPageCount;
                dr["URL"] = strURL + strSUB_NUM + ".html";
                dr["SubNum"] = strSUB_NUM;
                dr["IsDownLoad"] = 0;
                dr["Status"] = "未下載";
                dt.Rows.Add(dr);
            }

            GetliInfo2(strURL_O, NewContents2, ref dt);

        }

        /// <summary>
        /// 取得關鍵字串後的所有內容
        /// </summary>
        /// <param name="Contents">原內容</param>
        /// <param name="NewContents">找到關鍵字之後下半部的內容</param>
        /// <param name="key">關鍵字串</param>
        private void GetKey(string Contents, out string NewContents, string key)
        {
            NewContents = Contents.Substring(Contents.ToUpper().IndexOf(key.ToUpper()) + key.Length);

        }

        /// <summary>
        /// 取得Key內範圍的值
        /// </summary>
        /// <param name="Contents">原內容</param>
        /// <param name="Value">取出的值</param>
        /// <param name="keyS">開始的關鍵字串</param>
        /// <param name="keyE">結束的關鍵字串</param>
        private void GetValue(string Contents, out string Value, string keyS, string keyE)
        {
            Value = "";
            if (Contents.ToUpper().IndexOf(keyS.ToUpper()) >= 0)
            {
                try
                {
                    int IndexS = Contents.ToUpper().IndexOf(keyS.ToUpper());
                    string Temp1 = Contents.Substring(IndexS + keyS.Length);//後半
                    int IndexE = Temp1.ToUpper().IndexOf(keyE.ToUpper());
                    string Temp2 = Temp1.Substring(0, IndexE);//內容
                    Value = Temp2;
                }
                catch (Exception ex)
                {
                    //throw;
                }

            }

        }

        /// <summary>
        /// 取得Key內範圍的值，並回傳取完之後剩下的內容
        /// </summary>
        /// <param name="Contents">原內容</param>
        /// <param name="NewContents">取完值之後下半部的內容</param>
        /// <param name="Value">取出的值</param>
        /// <param name="keyS">開始的關鍵字串</param>
        /// <param name="keyE">結束的關鍵字串</param>
        private void GetValue(string Contents, out string NewContents, out string Value, string keyS, string keyE)
        {
            Value = "";
            NewContents = "";
            if (Contents.ToUpper().IndexOf(keyS.ToUpper()) >= 0)
            {
                try
                {
                    int IndexS = Contents.ToUpper().IndexOf(keyS.ToUpper());
                    string Temp1 = Contents.Substring(IndexS + keyS.Length);//後半
                    int IndexE = Temp1.ToUpper().IndexOf(keyE.ToUpper());
                    string Temp2 = Temp1.Substring(0, IndexE);//內容
                    Value = Temp2;

                    int IndexE2 = Contents.ToUpper().IndexOf(keyE.ToUpper(), IndexS) + keyE.Length;
                    if (IndexE2 < Contents.Length - 1)
                        NewContents = Contents.Substring(IndexE2);
                    else
                        NewContents = "";
                }
                catch (Exception ex)
                {
                    //throw;
                }

            }

        }

        List<WorkOption> listWorkOption = new List<WorkOption>();
        private void btnGetBySetting_Click(object sender, EventArgs e)
        {
            if (gvList.DataSource == null)
            {
                MessageBox.Show("請先取得項目!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                txtURL.Focus();
                txtURL.SelectAll();
                return;
            }
            GetAllWorkNew();
        }


        Random random = new Random();
        private void GetAllWorkNew()
        {
            #region 取得存檔資料夾路徑
            string DirPath = "";
            if (chkIsUseDefaultSavePath.Checked)//使用預設儲存路徑
            {
                if (txtDirPath.Text.Trim() == "")//無設定時
                {
                    MessageBox.Show("預設存檔路徑錯誤!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                    txtDirPath.Focus();
                    txtDirPath.SelectAll();
                    return;
                }
                else
                {
                    if (!Directory.Exists(txtDirPath.Text.Trim()))
                    {
                        try
                        {
                            Directory.CreateDirectory(txtDirPath.Text.Trim());
                            DirPath = txtDirPath.Text.Trim();
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("預設存檔路徑錯誤!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                            return;
                        }
                    }
                    else
                        DirPath = txtDirPath.Text.Trim();
                }
            }
            else
            {
                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                    DirPath = folderBrowserDialog1.SelectedPath;
                else
                    return;
            }
            if (DirPath.LastIndexOf("\\") != DirPath.Length - 1)
                DirPath += "\\";
            #endregion

            #region 取得漫畫名稱
            string MainName = "";
            if (txtMainName.Text.Trim() == "")
            {
                MessageBox.Show("漫畫名稱錯誤!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                txtMainName.Focus();
                txtMainName.SelectAll();
                return;
            }
            else
                MainName = txtMainName.Text.Trim();

            if (MainName.LastIndexOf("\\") != MainName.Length - 1)
                MainName += "\\";
            #endregion


            #region 內頁重覆的處理方式
            int SaveProcessType = 0;
            if (rdbContinue.Checked)
                SaveProcessType = 0;
            if (rdbCover.Checked)
                SaveProcessType = 1;
            if (rdbRename.Checked)
                SaveProcessType = 2;
            #endregion

            #region 取得下載清單
            DataTable dt = (DataTable)gvList.DataSource;
            //判斷選取
            var vSEL = from t in dt.AsEnumerable()
                       where t["SEL"].ToString().ToUpper() == "TRUE"
                       select t;
            if (vSEL != null && vSEL.Count() > 0)
                dt = vSEL.CopyToDataTable();
            else
                dt = dt.Clone();
            #endregion
            P1 = "#p=";


            for (int i = 0; i < dt.Rows.Count; i++)//清單內每一筆
            {
                DataRow dr = dt.Rows[i];
                int PageCount;
                if (!int.TryParse(dr["PageCount"].ToString().Trim(), out PageCount)) continue;
                string Name = dr["Name"].ToString().Trim();
                string URL = dr["URL"].ToString().Trim();
                AddWork(i, URL, DirPath, MainName, Name, PageCount, ScanWidth, ScanHeight, SaveProcessType);
            }
            ShowWorkList();
            //CheckWorkAndStartWork();//檢查是否有未執行的工作並執行工作ShowWorkList(取完所有清單後)

        }
        int backColorR = 0;
        int backColorG = 255;
        int backColorB = 0;

        private void AddWork(int Index, string URL, string DirPath, string MainName, string Name, int PageCount, int Width, int Height, int SaveProcessType)
        {
            string ID = DateTime.Now.ToString("MMddHHmmssfff") + random.Next(100, 999);

            //加入list
            WorkOption workoption = new WorkOption(new Thread(DoDownLoadWorkobj), Index + 1, ID, 0, 0, null, URL, 1, PageCount, DirPath, MainName, string.Format("{0}[{1}p]", Name, PageCount), "就緒", Width, Height, SaveProcessType, 0, pbOne, 20);
            listWorkOption.Add(workoption);

        }
        private delegate void myUICallBack(bool obj);

        private async void ShowWorkList(bool IsCheckWork = true)//顯示項目
        {
            if (this.InvokeRequired)//對跨執行緒做處理
            {
                myUICallBack myUpdate = new myUICallBack(ShowWorkList);
                this.Invoke(myUpdate, IsCheckWork);
            }
            else
            {


                //先重置
                FLPWorkList.Controls.Clear();
                FLPWorkList.AutoScroll = true;
                //FLPWorkList.BackColor = Color.Khaki;
                for (int i = 0; i < listWorkOption.Count; i++)
                {
                    WorkOption workoption = listWorkOption[i];
                    if (!chkShowComplete.Checked && workoption.Status == 2)
                        continue;
                    #region 顯示項目
                    workoption.FLP = new FlowLayoutPanel();
                    workoption.FLP.Width = 1095;//總寬度
                    workoption.FLP.Height = 30;
                    workoption.FLP.BackColor = Color.Khaki;
                    workoption.FLP.Dock = DockStyle.None;

                    //選取
                    workoption.chkSEL = new CheckBox();
                    workoption.chkSEL.Text = "";
                    workoption.chkSEL.AutoSize = false;
                    workoption.chkSEL.Size = new Size(23, 23);
                    workoption.chkSEL.TextAlign = ContentAlignment.MiddleLeft;
                    workoption.chkSEL.Padding = new Padding(0);
                    workoption.chkSEL.Margin = new Padding(0);
                    workoption.FLP.Controls.Add(workoption.chkSEL);

                    //序號
                    workoption.labNo = new Label();
                    workoption.labNo.Text = workoption.NO.ToString();
                    workoption.labNo.AutoSize = false;
                    workoption.labNo.Size = new Size(40, 23);
                    workoption.labNo.TextAlign = ContentAlignment.MiddleLeft;
                    workoption.labNo.Padding = new Padding(0);
                    workoption.labNo.Margin = new Padding(0);
                    workoption.FLP.Controls.Add(workoption.labNo);

                    //狀態
                    workoption.labStatus = new Label();
                    int Status = workoption.Status;
                    workoption.labStatus.Text = Status == 0 ? "未執行" : Status == 1 ? "進行中" : Status == 2 ? "已完成" : Status == 3 ? "暫停中" : Status == 4 ? "錯誤停止" : "不明";
                    workoption.labStatus.AutoSize = false;
                    workoption.labStatus.Size = new Size(60, 23);
                    workoption.labStatus.TextAlign = ContentAlignment.MiddleCenter;
                    workoption.labStatus.Padding = new Padding(0);
                    workoption.labStatus.Margin = new Padding(0);
                    workoption.FLP.Controls.Add(workoption.labStatus);

                    //卷名及開啟資料夾
                    workoption.btnOpenDir = new Button();
                    string MainName_ = workoption.MainName;
                    if (MainName_.Length > 5)//長度過長 截掉
                        MainName_ = MainName_.Substring(0, 3) + "...";
                    workoption.btnOpenDir.Text = string.Format("[{0}-{1}]", MainName_, workoption.Name);
                    workoption.btnOpenDir.Tag = workoption.ID;
                    workoption.btnOpenDir.AutoSize = false;
                    workoption.btnOpenDir.Size = new Size(200, 23);
                    workoption.btnOpenDir.TextAlign = ContentAlignment.MiddleCenter;
                    workoption.btnOpenDir.Padding = new Padding(0);
                    workoption.btnOpenDir.Margin = new Padding(0);
                    workoption.btnOpenDir.Click += btnOpenDir_Click;
                    workoption.FLP.Controls.Add(workoption.btnOpenDir);

                    //執行數量狀況
                    workoption.labWorkCount = new Label();
                    string WorkCount = string.Format("{0}/{1}", workoption.PageIndex, workoption.PageCount); //執行數量狀況
                    workoption.labWorkCount.Text = WorkCount;
                    workoption.labWorkCount.AutoSize = false;
                    workoption.labWorkCount.Size = new Size(80, 23);
                    workoption.labWorkCount.TextAlign = ContentAlignment.MiddleRight;
                    workoption.labWorkCount.Padding = new Padding(0);
                    workoption.labWorkCount.Margin = new Padding(0);
                    workoption.FLP.Controls.Add(workoption.labWorkCount);

                    //訊息
                    workoption.labStatus2 = new Label();
                    workoption.labStatus2.Text = workoption.Message;//訊息
                    workoption.labStatus2.AutoSize = false;
                    workoption.labStatus2.Size = new Size(150, 23);
                    workoption.labStatus2.TextAlign = ContentAlignment.MiddleCenter;
                    workoption.labStatus2.Padding = new Padding(0);
                    workoption.labStatus2.Margin = new Padding(0);
                    workoption.labStatus2.ForeColor = Color.Red;
                    workoption.FLP.Controls.Add(workoption.labStatus2);

                    workoption.btnMain = new Button();
                    workoption.btnMain.Size = new Size(60, 23);
                    workoption.btnMain.Click += btnMain_Click;
                    workoption.btnMain.Tag = workoption.ID;
                    workoption.FLP.Controls.Add(workoption.btnMain);

                    workoption.btnName = new Button();
                    workoption.btnName.Size = new Size(60, 23);
                    workoption.btnName.Click += btnName_Click;
                    workoption.btnName.Tag = workoption.ID;
                    workoption.FLP.Controls.Add(workoption.btnName);

                    workoption.btnSkip = new Button();
                    workoption.btnSkip.Size = new Size(46, 23);
                    workoption.btnSkip.Click += btnSkip_Click;
                    workoption.btnSkip.Tag = workoption.ID;
                    workoption.FLP.Controls.Add(workoption.btnSkip);

                    workoption.btnStop = new Button();
                    workoption.btnStop.Text = Status == 0 ? "暫停" : Status == 1 ? "暫停" : Status == 3 ? "繼續" : "--";
                    workoption.btnStop.Size = new Size(46, 23);
                    workoption.btnStop.Click += btnStop_Click;
                    workoption.btnStop.Tag = workoption.ID;
                    workoption.FLP.Controls.Add(workoption.btnStop);

                    workoption.btnDel = new Button();
                    workoption.btnDel.Text = "刪除";
                    workoption.btnDel.Size = new Size(46, 23);
                    workoption.btnDel.Tag = workoption.ID;
                    workoption.btnDel.Click += btnDelete_Click;
                    workoption.FLP.Controls.Add(workoption.btnDel);

                    FLPWorkList.Controls.Add(workoption.FLP);
                    #endregion
                }

                if (IsCheckWork)
                {
                    CheckWorkAndStartWork();//檢查所有工作並執行

                }
            }
        }
  
        private void CheckWorkAndStartWork()//檢查是否有未執行的工作並執行工作ShowWorkList
        {

                try
                {
                    //搜尋是否有執行中工作
                    var v = from WorkOption t in listWorkOption
                            where t.IsWork == 1
                            select t;

                    //可設定同時執行數
                    int intWorkCount = CheckWorkCount();
                    if (v.Count() < intWorkCount)//若工作在執行數量少於設定值時
                    {
                        int WorkingCNT = 0;//累計已執行數
                        for (int i = 0; i < listWorkOption.Count; i++)
                        {
                            if (listWorkOption[i].IsWork == 1)//計算已執行數
                            {
                                WorkingCNT++;
                                continue;
                            }

                            if (listWorkOption[i].IsWork == 0 && WorkingCNT < intWorkCount)//未執行的 並且執行數量少於設定值時 && listWorkOption[i].Status==0
                            {

                                switch (listWorkOption[i].Status)
                                {
                                    case 0://初始

                                        if (false)
                                        {
                                            DoDownLoadWork(listWorkOption[i]);
                                        }
                                        else //(rdbProcess1.Checked)
                                        {

                                            if (!listWorkOption[i].thread.IsAlive)
                                            {
                                                listWorkOption[i].thread = new Thread(DoDownLoadWorkobj);
                                            }

                                            listWorkOption[i].thread.IsBackground = true;//設為背景執行~才可刪除
                                                                                         //listWorkOption[i].thread.SetApartmentState(ApartmentState.MTA);
                                            listWorkOption[i].IsWork = 1;
                                            listWorkOption[i].Status = 1;
                                            listWorkOption[i].thread.Start(new object[] { listWorkOption[i] });//執行 執行緒
                                                                                                               //mStart.WaitOne();
                                        }
                                        WorkingCNT++;
                                        Thread.Sleep(1);
                                        break;
                                    case 1://執行中
                                        break;
                                    case 2://已完成
                                        try
                                        {
                                            if (listWorkOption[i].WB != null)
                                            {
                                                listWorkOption[i].WB.Stop();
                                                listWorkOption[i].WB.Dispose();
                                            }
                                            if (listWorkOption[i].thread != null && listWorkOption[i].thread.IsAlive)
                                            {
                                                listWorkOption[i].thread.Abort();
                                                listWorkOption[i].thread.Join();
                                            }
                                        }
                                        catch (Exception) { }
                                        break;
                                    case 3://暫停中
                                        try
                                        {
                                            if (listWorkOption[i].WB != null)
                                            {
                                                listWorkOption[i].WB.Stop();
                                                listWorkOption[i].WB.Dispose();
                                            }
                                            if (listWorkOption[i].thread != null && listWorkOption[i].thread.IsAlive)
                                            {
                                                listWorkOption[i].thread.Abort();
                                                listWorkOption[i].thread.Join();
                                            }
                                        }
                                        catch (Exception) { }
                                        break;
                                    case 4://錯誤停止
                                        try
                                        {
                                            if (listWorkOption[i].WB != null)
                                            {
                                                listWorkOption[i].WB.Stop();
                                                listWorkOption[i].WB.Dispose();
                                            }
                                            if (listWorkOption[i].thread != null && listWorkOption[i].thread.IsAlive)
                                            {
                                                listWorkOption[i].thread.Abort();
                                                listWorkOption[i].thread.Join();
                                            }
                                        }
                                        catch (Exception) { }
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }


                    //檢查是否所有工作皆完成 0初始 1執行中 2已完成 3暫停中 4錯誤停止
                    var v2 = from WorkOption t in listWorkOption
                             where t.Status == 2
                             select t;
                    var v3 = from WorkOption t in listWorkOption
                             where t.Status == 3
                             select t;
                    var v4 = from WorkOption t in listWorkOption
                             where t.Status == 4
                             select t;
                    labStatus.Text = string.Format("目前{0}個工作執行中，已完成{1}個工作，未執行{2}個工作，暫停{3}個工作，錯誤{4}個工作",
                                                    v.Count(),
                                                    v2.Count(),
                                                    listWorkOption.Count - v.Count() - v2.Count() - v3.Count() - v4.Count(),
                                                    v3.Count(),
                                                    v4.Count());//狀態列
                    pbAll.Maximum = listWorkOption.Count;
                    pbAll.Value = v2.Count();


                    if (v2 != null && v2.Count() > 0)
                    {
                        if (v2.Count() == listWorkOption.Count)
                        {
                            labStatus.Text = "工作完成";
                            if (chkFinishAlert.Checked)//工作完成提示
                                MessageBox.Show("所有工作都執行完畢!", "完成", MessageBoxButtons.OK, MessageBoxIcon.None);//狀態列


                            foreach (WorkOption item in v2)
                            {
                                if (item.thread == null || !item.thread.IsAlive) continue;
                                if (item.WB != null)
                                {
                                    item.WB.Stop();
                                    item.WB.Dispose();
                                }
                                item.thread.Abort();
                                item.thread.Join();
                            }
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {

                }
  

        }

        #region 按鈕
        public void btnOpenDir_Click(object sender, EventArgs e)//開啟存檔資料夾
        {
            try
            {
                string ID = ((System.Windows.Forms.ButtonBase)(sender)).Tag.ToString();
                var V = (from WorkOption t in listWorkOption
                         where t.ID == ID
                         select t).Take(1);
                WorkOption workoption = (WorkOption)V.ToList()[0];


                string SubDirPath = workoption.DirPath + workoption.MainName + workoption.Name; ;
                if (!Directory.Exists(SubDirPath))
                    Directory.CreateDirectory(SubDirPath);
                if (SubDirPath.Trim() != "")
                    System.Diagnostics.Process.Start(SubDirPath);

            }
            catch (Exception ex)
            {

                throw;
            }
        }
        private void btnMain_Click(object sender, EventArgs e)//
        {

            //搜尋是否有執行中工作
            Button btn = (Button)sender;
            //用ID值去找工作
            var v = from WorkOption t in listWorkOption
                    where t.ID.Trim() == btn.Tag.ToString().Trim()
                    select t;

            if (v != null && v.Count() > 0)
            {
                WorkOption workoption = (WorkOption)v.ToList()[0];

                try
                {
                    string URL = workoption.URL.Trim().Substring(0, workoption.URL.Trim().LastIndexOf("/"));
                    System.Diagnostics.Process.Start("Chrome.EXE", URL);
                }
                catch (Exception)
                {

                    System.Diagnostics.Process.Start("Internet Explorer.EXE", workoption.URL.ToString());
                }
            }

        }
        private void btnName_Click(object sender, EventArgs e)//
        {

            //搜尋是否有執行中工作
            Button btn = (Button)sender;
            //用ID值去找工作
            var v = from WorkOption t in listWorkOption
                    where t.ID.Trim() == btn.Tag.ToString().Trim()
                    select t;

            if (v != null && v.Count() > 0)
            {
                WorkOption workoption = (WorkOption)v.ToList()[0];

                string strFinalURL = string.Format("{0}{1}{2}", workoption.URL.Trim(), "#p=", workoption.PageIndex);
                try
                {
                    System.Diagnostics.Process.Start("Chrome.EXE", strFinalURL);
                }
                catch (Exception)
                {

                    System.Diagnostics.Process.Start("Internet Explorer.EXE", strFinalURL);
                }
            }


        }
        private void btnSkip_Click(object sender, EventArgs e)//
        {

            //搜尋是否有執行中工作
            Button btn = (Button)sender;
            //用ID值去找工作
            var v = from WorkOption t in listWorkOption
                    where t.ID.Trim() == btn.Tag.ToString().Trim()
                    select t;

            if (v != null && v.Count() > 0)
            {
                WorkOption workoption = (WorkOption)v.ToList()[0];

                if (workoption.Status == 1)//進行中 換成 完成
                {
                    workoption.IsWork = 0;//
                    workoption.Status = 2;//完成
                    if (workoption.thread != null && workoption.thread.IsAlive)
                        workoption.thread.Interrupt();

                }
                else if (workoption.Status == 2)//已完成 不動作
                {
                    MessageBox.Show("工作已完成無法暫停!", "警告", MessageBoxButtons.OK, MessageBoxIcon.None);
                    return;
                }
                else if (workoption.Status == 0)//未執行 換成 暫停
                {
                    workoption.IsWork = 0;//
                    workoption.Status = 2;//完成
                }
                else if (workoption.Status == 3)//暫停中 換成 進行中
                {
                    workoption.IsWork = 0;
                    workoption.Status = 2;
                }
                workoption.Message = "跳過此工作";
            }

            CheckWorkAndStartWork();//(暫停/繼續工作)
        }
        private void btnStop_Click(object sender, EventArgs e)//暫停/繼續工作
        {

            //搜尋是否有執行中工作
            Button btn = (Button)sender;
            //用ID值去找工作
            var v = from WorkOption t in listWorkOption
                    where t.ID.Trim() == btn.Tag.ToString().Trim()
                    select t;

            if (v != null && v.Count() > 0)
            {
                WorkOption workoption = (WorkOption)v.ToList()[0];

                if (workoption.Status == 1)//進行中 換成 暫停
                {
                    workoption.IsWork = 2;//暫停
                    workoption.Status = 3;//暫停
                    if (workoption.thread != null && workoption.thread.IsAlive)
                        workoption.thread.Interrupt();

                    workoption.Message += "-暫停";

                }
                else if (workoption.Status == 2)//已完成 不動作
                {
                    MessageBox.Show("工作已完成無法暫停!", "警告", MessageBoxButtons.OK, MessageBoxIcon.None);
                    return;
                }
                else if (workoption.Status == 0)//未執行 換成 暫停
                {
                    workoption.IsWork = 2;//暫停
                    workoption.Status = 3;//暫停
                    workoption.Message += "-暫停";
                }
                else if (workoption.Status == 3)//暫停中 換成 進行中
                {
                    workoption.IsWork = 0;
                    workoption.Status = 0;
                    workoption.Message = workoption.Message.Replace("-暫停", "");
                }
            }

            CheckWorkAndStartWork();//(暫停/繼續工作)
        }
        private void btnDelete_Click(object sender, EventArgs e)//刪除工作
        {
            try
            {
                Button btn = (Button)sender;
                DialogResult result = MessageBox.Show("確定刪除工作?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                if (result == System.Windows.Forms.DialogResult.Cancel)//取消
                    return;

                //用ID值去找工作
                var v = from WorkOption t in listWorkOption
                        where t.ID.Trim() == btn.Tag.ToString().Trim()
                        select t;
                var vRemove = from WorkOption t in listWorkOption
                              where t.ID.Trim() != btn.Tag.ToString().Trim()
                              select t;
                if (v != null && v.Count() > 0)
                {
                    WorkOption workoption = (WorkOption)v.ToList()[0];
                    int Status = workoption.Status;
                    if (Status == 1)//進行中
                    {
                        DialogResult result2 = MessageBox.Show("工作在進行中確定要刪除?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                        if (result2 == System.Windows.Forms.DialogResult.Cancel)//取消
                            return;

                        try
                        {
                            if (workoption.thread != null && workoption.thread.IsAlive)
                            {
                                if (workoption.WB != null)
                                {
                                    workoption.WB.Stop();
                                    workoption.WB.Dispose();
                                }
                                workoption.thread.Interrupt();
                                workoption.thread.Abort();
                                workoption.thread.Join();//用來等候Thread投擲出"使Thread停止的exception
                                listWorkOption.Remove(workoption);//從清單內刪除
                            }
                            Thread.Sleep(1000);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("工作執行中，錯誤無法刪除!", "警告", MessageBoxButtons.OK, MessageBoxIcon.None);
                        }
                    }
                    //else if (Sataus == 2)//已完成
                    //{
                    //    MessageBox.Show("工作已完成無法刪除!", "警告", MessageBoxButtons.OK, MessageBoxIcon.None);
                    //    return;
                    //}
                    else//其他狀態 刪除
                        listWorkOption.Remove(workoption);//從清單內刪除
                }
                listWorkOption = new List<WorkOption>(vRemove.ToList());
                CheckWorkAndStartWork();//執行下一份工作
                FLPWorkList.Controls.Clear();
                ShowWorkList();//顯示項目

                MessageBox.Show("刪除工作成功!", "成功", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
            catch (Exception ex)
            {
                if (ex.Message.IndexOf("集合") > -1)
                {
                    ShowWorkList();//刪除工作後顯示新清單
                    return;
                }
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        private int CheckWorkCount()//取得同時執行工作數量
        {
            if (cmbWorkCount.SelectedItem.ToString().Trim() == "") return 1;//空值回傳1

            if (cmbWorkCount.SelectedItem.ToString().Trim() == "全部")//全部回傳總數量
            {
                return listWorkOption.Count;
            }
            else if (cmbWorkCount.SelectedItem.ToString().Trim() == "自訂")//自訂回傳設定值
            {
                int Count;
                if (int.TryParse(txtWorkCount.Text, out Count))
                    return Count;
                else
                    return 1;
            }
            else
            {
                int Count;
                if (int.TryParse(cmbWorkCount.SelectedItem.ToString().Trim(), out Count))
                    return Count;
                else
                    return 1;
            }
        }


        private delegate void myDoDownLoadWorkobj(object obj);
        private void DoDownLoadWorkobj(object obj)
        {
            if (this.InvokeRequired)//對跨執行緒做處理
            {

                myDoDownLoadWorkobj mydodownloadworkobj = new myDoDownLoadWorkobj(DoDownLoadWorkobj);
                this.Invoke(mydodownloadworkobj, obj);

            }
            else
            {
                WorkOption workoption = (WorkOption)((object[])obj)[0];
                DoDownLoadWork(workoption);
            }
        }

        private delegate void myDoDownLoadWorkk(WorkOption workoption);

        private async void DoDownLoadWork(WorkOption workoption)//由內頁抓取連結下載
        {
            //if (this.InvokeRequired)//對跨執行緒做處理
            //{
            //    myDoDownLoadWorkk mydodownloadworkk = new myDoDownLoadWorkk(DoDownLoadWork);
            //    this.Invoke(mydodownloadworkk, workoption);
            //}
            //else
            //{

            //主頁
            try
            {
                //若是暫停中
                #region 若是暫停中
                if (workoption.Status == 3)
                {

                    if (workoption.thread != null && workoption.thread.IsAlive)
                    {
                        workoption.thread.Interrupt();//插斷執行緒
                        workoption.thread.Abort();//結束執行緒
                        workoption.thread.Join();//封鎖執行緒
                    }
                    return;
                }
                #endregion
                string strFinalURL = string.Format("{0}{1}{2}", workoption.URL.Trim(), "#p=", workoption.PageIndex);
                string SubDirPath = workoption.DirPath + workoption.MainName + workoption.Name;//DirPath + dr["Name"].ToString().Trim();
                string SubDirPath_ZIP = workoption.DirPath + workoption.MainName + workoption.Name + ".zip";
                string SubDirPath_ZIP2 = workoption.DirPath + workoption.MainName + workoption.Name + ".rar";
                string SubDirPath_ZIP3 = workoption.DirPath + workoption.MainName + workoption.Name + ".7z";

                #region 判斷是否已下載
                bool IsDO = true;
                if (Directory.Exists(SubDirPath))
                {
                    //有壓縮檔時 不下載
                    if (File.Exists(SubDirPath_ZIP) || File.Exists(SubDirPath_ZIP2) || File.Exists(SubDirPath_ZIP3))
                    {
                        IsDO = false;
                    }
                    else
                    {
                        int StartPageIndex = workoption.PageIndex;
                        for (int i = StartPageIndex; i <= workoption.PageCount; i++)
                        {
                            IsDO = CheckPageExist(workoption, SubDirPath);
                            if (!IsDO)
                            {
                                workoption.PageIndex++;
                                continue;
                            }
                            else
                                break;

                        }
                        //已完成工作時
                        #region 已完成工作時
                        if ((workoption.PageCount == 0 || workoption.PageIndex > workoption.PageCount) && !IsDO)
                        {
                            workoption.PageIndex = workoption.PageCount;

                            //將執行序停止並釋放
                            if (workoption.thread != null && workoption.thread.IsAlive)
                            {
                                workoption.thread.Abort();
                                workoption.thread.Join();
                            }
                            workoption.IsWork = 0;//無執行
                            workoption.Status = 2;//已完成
                            workoption.Message = "完成下載";
                            if (!chkShowComplete.Checked)//如果不Show完成下載
                                workoption.FLP.Visible = false;

                            if (chkRecord.Checked)
                            {
                                string strNUM = "";
                                string strSUB_NUM = "";
                                string P = "";
                                string strURL = "";
                                GetNumURL(workoption.URL, out strURL, out strNUM, out strSUB_NUM, out P);
                                AddRecord(strURL, strNUM, workoption.Name, strSUB_NUM);
                            }
                            CheckWorkAndStartWork();//更新狀態(完成下載)
                            return;
                        }
                        #endregion

                    }
                }
                #endregion



                workoption.SleepMilliseconds = 0;
                if (IsDO)//要執行下載時
                {
                    workoption.Message = "下載第" + workoption.PageIndex + "頁";
                    //透過get_manhuagui.exe
                    string Comm = string.Format("\"{0}\" \"{1}\"", Directory.GetCurrentDirectory() + "\\get_manhuagui.exe", strFinalURL);

                    string Result = RunProcess(Comm);//取得結果

                    int Index00 = Result.LastIndexOf("#");//內容有#時
                    if (Index00 > -1)
                    {
                        //取得每一個內頁圖片連結
                        Result = Result.Substring(Index00 + 1);

                        string[] Links = Result.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                        //下載每筆內頁
                        for (int i = 0; i < Links.Length; i++)
                        {
                            //跳過先前已下載的頁數
                            if (i + 1 < workoption.PageIndex)
                                continue;
                            if (!CheckPageExist(workoption, SubDirPath))//已有該檔跳過
                                continue;
                            string strSRC = Links[i].Trim();
                            //非連結，跳過
                            if (strSRC.ToUpper().IndexOf("HTTP") < 0)
                                continue;
                            workoption.Message = string.Format("下載第{0}頁", workoption.PageIndex+1);

                            //CheckWorkAndStartWork();//更新狀態
                            //直接下載
                            DownLoadLink(ref workoption, strSRC);

                            workoption.RetryCnt = 0;//重試歸零
                                                    //是否需要延遲下載，是的話隨機暫停時數
                            if (ckbIsSleep.Checked)
                            {
                                Random random = new Random();
                                int Sleep_S = 10, Sleep_E = 12;
                                if (!int.TryParse(cmbSleep_S.Text.Trim(), out Sleep_S))
                                    Sleep_S = 10;
                                if (!int.TryParse(cmbSleep_E.Text.Trim(), out Sleep_E))
                                    Sleep_E = 12;
                                if (Sleep_E < Sleep_S)
                                {
                                    Sleep_S = 10;
                                    Sleep_E = 12;
                                }
                                int intSleep = random.Next(Sleep_S, Sleep_E);
                                await Task.Delay(intSleep*1000);
                                workoption.PageIndex++;//頁碼增加
                            }//是否需要延遲下載 結束

                            
                            break;
                        }//下載每筆內頁 結束

                        //已完成工作時
                        #region 已完成工作時
                        if ((workoption.PageCount == 0 || workoption.PageIndex > workoption.PageCount))
                        {
                            workoption.PageIndex = workoption.PageCount;

                            //將執行序停止並釋放
                            if (workoption.thread != null && workoption.thread.IsAlive)
                            {
                                workoption.thread.Abort();
                                workoption.thread.Join();
                            }
                            workoption.IsWork = 0;//無執行
                            workoption.Status = 2;//已完成
                            workoption.Message = "完成下載";
                            if (!chkShowComplete.Checked)//如果不Show完成下載
                                workoption.FLP.Visible = false;

                            if (chkRecord.Checked)
                            {
                                string strNUM = "";
                                string strSUB_NUM = "";
                                string P = "";
                                string strURL = "";
                                GetNumURL(workoption.URL, out strURL, out strNUM, out strSUB_NUM, out P);
                                AddRecord(strURL, strNUM, workoption.Name, strSUB_NUM);
                            }
                            CheckWorkAndStartWork();//更新狀態(完成下載)
                            return;
                        }

                        //未完成工作時，報錯，進行下一個工作
                        //CheckWorkAndStartWork();//更新狀態
                        DoDownLoadWork(workoption);//進行下一頁工作
                        //workoption.thread.Start(new object[] { workoption });//執行 執行緒
                        //ShowWorkList(false);
                        //DoDownLoadWorkobj(new object[] { workoption });
                        #endregion
                    }///內容有#時 結束
                    else//內容沒有有#時 錯誤
                    {
                        //將執行序停止並釋放
                        if (workoption.thread != null && workoption.thread.IsAlive)
                        {
                            workoption.thread.Abort();
                            workoption.thread.Join();
                        }
                        workoption.IsWork = 0;//無執行
                        workoption.Status = 4;//錯誤停止
                        workoption.Message = "停止，get_manhuagui.exe錯誤";
                        CheckWorkAndStartWork();//更新狀態(下載內頁圖片 Exception錯誤)
                    }
                }//要執行下載時 結束

            }
            catch (Exception ex)
            {
                //將執行序停止並釋放
                if (workoption.thread != null && workoption.thread.IsAlive)
                {
                    workoption.thread.Abort();
                    workoption.thread.Join();
                }
                workoption.IsWork = 0;//無執行
                workoption.Status = 4;//錯誤停止
                workoption.Message = "停止，錯誤01";
                File.WriteAllText("D:\\TEMP\\Log.txt",ex.Message+ex.StackTrace);
                CheckWorkAndStartWork();//更新狀態(下載內頁圖片 Exception錯誤)
            }
            //}//跨執行緒結束
        }

        private static bool CheckPageExist(WorkOption workoption, string SubDirPath)
        {
            bool IsDO = true;
            DirectoryInfo dirinfo = new DirectoryInfo(SubDirPath);
            if (!Directory.Exists(SubDirPath))//沒有該資料夾時
            {
                IsDO = true;
                return IsDO;
            }
            FileInfo[] files = dirinfo.GetFiles();
            string PageIndexTemp = workoption.PageIndex.ToString();
            int IndexFile = files.ToList().FindIndex(f => f.Name.Substring(0, f.Name.LastIndexOf(".")) == PageIndexTemp.PadLeft(3, '0').Trim());
            if (IndexFile >= 0)//已有該檔時
            {
                switch (workoption.SaveProcessType)
                {
                    case 0://跳過
                        IsDO = false;
                        break;
                    case 1://覆蓋
                        IsDO = true;
                        break;
                    case 2://ReName
                        IsDO = true;
                        break;
                    default:
                        IsDO = false;
                        break;
                }

            }

            return IsDO;
        }

        //直接使用正確連結下載
        private void DownLoadLink(ref WorkOption workoption,string strSRC)
        {
            try
            {
                //下載
                string FE = ".jpg";
                int Index01 = strSRC.LastIndexOf(".");
                int Index02 = strSRC.IndexOf("?", Index01);
                if (Index02 > -1)
                    FE = strSRC.Substring(Index01, Index02 - Index01);//副檔名
                else
                    FE = strSRC.Substring(Index01);//副檔名
                if (FE.Length > 4)
                    FE = ".jpg";

                string SubDirPath = workoption.DirPath + workoption.MainName + workoption.Name;
                string FilePath = SubDirPath + "\\" + workoption.PageIndex.ToString().PadLeft(3, '0') + FE;//組合檔名
                if (!Directory.Exists(SubDirPath + "\\"))
                    Directory.CreateDirectory(SubDirPath + "\\");

                string ErrorMessage = DoDownLoadPic_ByWebRequest(ref workoption, strSRC, FilePath, FE);
                if (ErrorMessage.Trim() != "")//有錯誤時
                {
                    if (workoption.RetryCnt <= 3)//重試
                    {

                        workoption.RetryCnt++;
                        DoDownLoadWork(workoption);//再試一次
                                                   //DoDownLoadWorkobj(workoption);//繼續下載※
                        return;
                    }
                    else
                    {
                        if (workoption.thread != null && workoption.thread.IsAlive)
                        {
                            //將執行序停止並釋放
                            workoption.thread.Abort();
                            workoption.thread.Join();
                        }
                        workoption.Message = "錯誤，重試" + workoption.RetryCnt;
                        workoption.IsWork = 0;
                        workoption.Status = 4;//錯誤
                        CheckWorkAndStartWork();//更新狀態(下載內頁圖片 錯誤，重試)，自動執行下一個工作
                        return;
                    }
                }
                    
            }
            catch (Exception ex)
            {
                
                //將執行序停止並釋放
                if (workoption.thread != null && workoption.thread.IsAlive)
                {
                    workoption.thread.Abort();
                    workoption.thread.Join();
                }
                workoption.IsWork = 0;//無執行
                workoption.Status = 4;//錯誤停止
                workoption.Message = "停止，錯誤02:" + ex.Message;
                CheckWorkAndStartWork();//更新狀態(下載內頁圖片 Exception錯誤)
                return;
            }
        }

        //舊方式-使用瀏覽器截圖(卡在Loading那邊，下載約10頁後會掛掉)
        private void DoDownLoadWork_Old(WorkOption workoption)//導至主頁 string SubDirPath, string strURL, ref int PageIndex, int PageEnd
        {
            //if (this.InvokeRequired)//對跨執行緒做處理
            //{
            //    myDoDownLoadWorkk mydodownloadworkk = new myDoDownLoadWorkk(DoDownLoadWork);
            //    this.Invoke(mydodownloadworkk, workoption);
            //}
            //else
            //{

            //主頁
            try
            {
                //若是暫停中
                if (workoption.Status == 3)
                {
                    if (workoption.WB != null)
                    {
                        workoption.WB.Stop();
                        workoption.WB.Dispose();
                    }
                    if (workoption.thread != null && workoption.thread.IsAlive)
                    {
                        workoption.thread.Interrupt();//插斷執行緒
                        workoption.thread.Abort();//結束執行緒
                        workoption.thread.Join();//封鎖執行緒
                    }
                    return;
                }
                string strFinalURL = string.Format("{0}{1}{2}", workoption.URL.Replace("http:","https:").Trim(), "#p=", workoption.PageIndex);
                string SubDirPath = workoption.DirPath + workoption.MainName + workoption.Name;//DirPath + dr["Name"].ToString().Trim();
                string SubDirPath_ZIP = workoption.DirPath + workoption.MainName + workoption.Name + ".zip";
                string SubDirPath_ZIP2 = workoption.DirPath + workoption.MainName + workoption.Name + ".rar";
                string SubDirPath_ZIP3 = workoption.DirPath + workoption.MainName + workoption.Name + ".7z";

                #region 判斷是否已下載
                bool IsDO = true;
                if (Directory.Exists(SubDirPath))
                {
                    //有壓縮檔時 不下載
                    if (File.Exists(SubDirPath_ZIP) || File.Exists(SubDirPath_ZIP2) || File.Exists(SubDirPath_ZIP3))
                    {
                        IsDO = false;
                    }
                    else
                    {
                        DirectoryInfo dirinfo = new DirectoryInfo(SubDirPath);
                        FileInfo[] files = dirinfo.GetFiles();
                        string PageIndexTemp = workoption.PageIndex.ToString();
                        int IndexFile = files.ToList().FindIndex(f => f.Name.Substring(0, f.Name.LastIndexOf(".")) == PageIndexTemp.PadLeft(3, '0').Trim());
                        if (IndexFile >= 0)//已有該檔時
                        {
                            switch (workoption.SaveProcessType)
                            {
                                case 0://跳過
                                    IsDO = false;
                                    break;
                                case 1://覆蓋
                                    IsDO = true;
                                    break;
                                case 2://ReName
                                    IsDO = true;
                                    break;
                                default:
                                    IsDO = false;
                                    break;
                            }

                        }
                    }
                }
                #endregion

                workoption.SleepMilliseconds = 0;
                if (IsDO)//要執行下載時
                {
                    //是否需要延遲下載
                    if (ckbIsSleep.Checked)
                    {
                        Random random = new Random();
                        int Sleep_S=10,Sleep_E = 12;
                        if (!int.TryParse(cmbSleep_S.Text.Trim(), out Sleep_S))
                            Sleep_S = 10;
                        if (!int.TryParse(cmbSleep_E.Text.Trim(), out Sleep_E))
                            Sleep_E = 12;
                        if (Sleep_E < Sleep_S)
                        {
                            Sleep_S = 10;
                            Sleep_E = 12;
                        }
                        int intSleep = random.Next(Sleep_S, Sleep_E);

                        workoption.SleepMilliseconds = intSleep * 1000;
                       
                    }

                    if (workoption.NavigateCNT > workoption.NavigateFlag)
                    {
                        ReNavigate(ref workoption);
                        return;
                    }
                    workoption.Message = "下載第" + workoption.PageIndex + "頁";

                    //if (workoption.PageIndex % 5 == 0)
                    ClearMemory();//記憶體回收

                    if (workoption.WB != null)
                    {
                        workoption.WB.Stop();
                        workoption.WB.Dispose();
                    }
                    workoption.WB = new WebBrowser();
                    IniWebBrowser();
                    workoption.WB.Width = workoption.Width;
                    workoption.WB.Height = workoption.Height;
                    workoption.WB.ScriptErrorsSuppressed = true;
                    workoption.WB.DocumentCompleted += wB_DocumentCompleted;
                    workoption.WB.Navigate(strFinalURL);//導頁
                    
                    int tryCount = 0;
                    loading(workoption.WB, ref tryCount);//等待完成載入
                    //IncludeScript(workoption.WB);//加入Script
                    workoption.NavigateCNT++;//導頁次數增加
                    Thread.Sleep(sleepSecond*1000);
                    DownLoadContent(ref workoption);//下載
                }
                else
                {
                    workoption.Message = "跳過第" + workoption.PageIndex + "頁";

                    //下一頁
                    workoption.PageIndex++;
                    //已完成工作時
                    if (workoption.PageCount == 0 || workoption.PageIndex > workoption.PageCount)
                    {
                        workoption.PageIndex = workoption.PageCount;
                        if (workoption.WB != null)
                        {
                            workoption.WB.Stop();
                            workoption.WB.Dispose();
                        }
                        //將執行序停止並釋放
                        if (workoption.thread != null && workoption.thread.IsAlive)
                        {
                            workoption.thread.Abort();
                            workoption.thread.Join();
                        }
                        workoption.IsWork = 0;//無執行
                        workoption.Status = 2;//已完成
                        workoption.Message = "完成下載";
                        if (!chkShowComplete.Checked)//如果不Show完成下載
                            workoption.FLP.Visible = false;

                        if (chkRecord.Checked)
                        {
                            string strNUM = "";
                            string strSUB_NUM = "";
                            string P = "";
                            string strURL = "";
                            GetNumURL(workoption.URL, out strURL, out strNUM, out strSUB_NUM, out P);
                            AddRecord(strURL, strNUM, workoption.Name, strSUB_NUM);
                        }
                        CheckWorkAndStartWork();//更新狀態(完成下載)
                        return;
                    }
                    else//繼續下載
                        DoDownLoadWork(workoption);//繼續下載
                        //DoDownLoadWorkobj(workoption);//繼續下載※
                }
            }
            catch (Exception ex)
            {

                DoDownLoadWork(workoption);//繼續下載
                //DoDownLoadWorkobj(workoption);//繼續下載※
            }
            //}
        }

        private void IniWebBrowser()
        {
            //程序啟動時運行
            if (!WBEmulator.IsBrowserEmulationSet())
            {
                WBEmulator.SetBrowserEmulationVersion();
            }
        }

        //加入Script
        private void IncludeScript(WebBrowser WB)
        {

            //<script src="https://unpkg.com/webp-hero@0.0.0-dev.21/dist-cjs/polyfills.js"></script>
            //<script src="https://unpkg.com/webp-hero@0.0.0-dev.21/dist-cjs/webp-hero.bundle.js"></script>
            //<script>
            //var webpMachine = new webpHero.WebpMachine();
            //webpMachine.polyfillDocument();
            //</script>


            //HtmlElement head = WB.Document.GetElementsByTagName("head")[0];
            //HtmlElement scriptEl = WB.Document.CreateElement("script");
            //IHTMLScriptElement element = (IHTMLScriptElement)scriptEl.DomElement;
            //element.text = "function sayHello() { ";
            //element.text += System.IO.File.ReadAllText(Environment.CurrentDirectory + "\\js\\polyfills.js");
            //element.text += System.IO.File.ReadAllText(Environment.CurrentDirectory + "\\js\\webp-hero.bundle.js");
            //element.text += "var webpMachine = new webpHero.WebpMachine();webpMachine.polyfillDocument(); }";
            //head.AppendChild(scriptEl);
            //WB.Document.InvokeScript("sayHello");
        }

        private void ReNavigate(ref WorkOption workoption)//重開執行緒 重新導頁
        {
            if(workoption.thread!=null && workoption.thread.IsAlive)
            {
                workoption.thread.Abort();
                workoption.thread.Join();
            }
            workoption.NavigateCNT = 0;
            workoption.thread = new Thread(DoDownLoadWorkobj);
            workoption.thread.IsBackground = true;//設為背景執行~才可刪除
            workoption.IsWork = 1;
            workoption.Status = 1;
            workoption.thread.Start(new object[] { workoption });//執行 執行緒
        }
        //下載內頁圖片
        private void DownLoadContent(ref WorkOption workoption)//string strURL, string SubDirPath, ref int PageIndex, int PageEnd)
        {
            HtmlDocument doc = workoption.WB.Document;
            #region 使用WebRequest下載
            for (int i = 0; i < doc.All.Count; i++)
            {
                //找到該Tag
                if (doc.All[i].TagName.Trim().ToUpper() == "DIV" && doc.All[i].Id != null && doc.All[i].Id.Trim().ToUpper() == "MANGABOX")
                {
                    string MANGABOX = doc.All[i].InnerHtml;//取得MANGABOX
                    #region 讀取圖片失败時
                    if (MANGABOX.IndexOf("失败") > -1)
                    {
                        sleepSecond++;
                        if (workoption.RetryCnt <= 3)//重試
                        {
                            if (workoption.RetryCnt > 2)//已重試3次時重置執行緒狀態 重新下載
                            {
                                workoption.NavigateCNT += 20;
                            }
                            workoption.RetryCnt++;
                            DoDownLoadWork(workoption);//再試一次
                                                       //DoDownLoadWorkobj(workoption);//繼續下載※
                            return;
                        }
                        else
                        {
                            workoption.WB.Stop();
                            workoption.WB.Dispose();
                            if (workoption.thread != null && workoption.thread.IsAlive)
                            {
                                //將執行序停止並釋放
                                workoption.thread.Abort();
                                workoption.thread.Join();
                            }
                            workoption.Message = "錯誤，重試" + workoption.RetryCnt;
                            workoption.IsWork = 0;
                            workoption.Status = 4;
                            CheckWorkAndStartWork();//更新狀態(下載內頁圖片 錯誤，重試)
                                                    //執行下一個工作

                            return;
                        }
                    }
                    #endregion

                    #region 取得圖片連結
                    string strSRC = "";
                    if (MANGABOX.Trim() != "")
                    {
                        GetValue(MANGABOX, out strSRC, "src=\"", "\"");
                    }
                    strSRC = HttpUtility.UrlDecode(strSRC);//取得SRC

                    #endregion

                    if (strSRC.Trim() == "")
                        continue;
                    sleepSecond = 3;//重設延持間
                    try
                    {
                        workoption.WB.Stop();
                        //下載
                        string FE = ".jpg";
                        int Index01 = strSRC.LastIndexOf(".");
                        int Index02 = strSRC.IndexOf("?", Index01);
                        if (Index02 > -1)
                            FE = strSRC.Substring(Index01, Index02 - Index01);//副檔名
                        else
                            FE = strSRC.Substring(Index01);//副檔名
                        if (FE.Length > 4)
                            FE = ".jpg";

                        string SubDirPath = workoption.DirPath + workoption.MainName + workoption.Name;
                        string FilePath = SubDirPath + "\\" + workoption.PageIndex.ToString().PadLeft(3, '0') + FE;//組合檔名
                        if (!Directory.Exists(SubDirPath + "\\"))
                            Directory.CreateDirectory(SubDirPath + "\\");

                        string ErrorMessage = DoDownLoadPic_ByWebRequest(ref workoption, strSRC, FilePath, FE);
                        if (ErrorMessage.Trim() != "")//有錯誤時
                        {
                            if (workoption.RetryCnt <= 3)//重試
                            {
                                if (workoption.RetryCnt > 2)//已重試3次時重置執行緒狀態 重新下載
                                {
                                    workoption.NavigateCNT += 20;
                                }
                                workoption.RetryCnt++;
                                DoDownLoadWork(workoption);//再試一次
                                                           //DoDownLoadWorkobj(workoption);//繼續下載※
                                return;
                            }
                            else
                            {
                                workoption.WB.Stop();
                                workoption.WB.Dispose();
                                if (workoption.thread != null && workoption.thread.IsAlive)
                                {
                                    //將執行序停止並釋放
                                    workoption.thread.Abort();
                                    workoption.thread.Join();
                                }
                                workoption.Message = "錯誤，重試" + workoption.RetryCnt;
                                workoption.IsWork = 0;
                                workoption.Status = 4;
                                CheckWorkAndStartWork();//更新狀態(下載內頁圖片 錯誤，重試)
                                                        //執行下一個工作

                                return;
                            }


                        }
                        workoption.RetryCnt = 0;
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("錯誤", ex.Message);
                        workoption.WB.Stop();
                        workoption.WB.Dispose();
                        //將執行序停止並釋放
                        if (workoption.thread != null && workoption.thread.IsAlive)
                        {
                            workoption.thread.Abort();
                            workoption.thread.Join();
                        }
                        workoption.IsWork = 0;//無執行
                        workoption.Status = 4;//錯誤停止
                        workoption.Message = "停止，錯誤03:" + ex.Message;
                        CheckWorkAndStartWork();//更新狀態(下載內頁圖片 Exception錯誤)
                        return;
                    }
                    //下一頁
                    workoption.PageIndex++;
                    if (workoption.PageCount == 0 || workoption.PageIndex > workoption.PageCount)
                    {
                        workoption.PageIndex = workoption.PageCount;
                        workoption.WB.Stop();
                        workoption.WB.Dispose();
                        //將執行序停止並釋放
                        if (workoption.thread != null && workoption.thread.IsAlive)
                        {
                            workoption.thread.Abort();
                            workoption.thread.Join();
                        }
                        workoption.IsWork = 0;//無執行
                        workoption.Status = 2;//已完成
                        workoption.Message = "完成下載";
                        if (!chkShowComplete.Checked)//如果不Show完成下載
                            workoption.FLP.Visible = false;
                        if (chkRecord.Checked)
                        {
                            string strNUM = "";
                            string strSUB_NUM = "";
                            string P = "";
                            string strURL = "";
                            GetNumURL(workoption.URL, out strURL, out strNUM, out strSUB_NUM, out P);
                            AddRecord(strURL, strNUM, workoption.Name, strSUB_NUM);
                        }
                        CheckWorkAndStartWork();//更新狀態(下載內頁圖片 完成下載)
                        return;
                    }

                    //ShowWorkList();//更新狀態
                    DoDownLoadWork(workoption);//繼續下載
                    //DoDownLoadWorkobj(workoption);//繼續下載※
                    return;
                }
            }
            //都沒找到IMG時
            if (workoption.RetryCnt <= 3)//重試
            {
                workoption.RetryCnt++;
                workoption.Message = "重試第" + workoption.PageIndex + "頁";
                DoDownLoadWork(workoption);//
                //DoDownLoadWorkobj(workoption);//繼續下載※
            }
            else
            {
                workoption.WB.Stop();
                workoption.WB.Dispose();
                //將執行序停止並釋放
                if (workoption.thread != null && workoption.thread.IsAlive)
                {
                    workoption.thread.Abort();
                    workoption.thread.Join();
                }
                workoption.IsWork = 0;//無執行
                workoption.Status = 4;//錯誤停止
                workoption.Message = "停止，錯誤04";
                CheckWorkAndStartWork();//更新狀態(下載內頁圖片 Exception錯誤)
                return;
            }
            return;
            #endregion


            #region 截圖下載
            for (int i = 0; i < doc.All.Count; i++)
            {
                //找到該Tag
                if (doc.All[i].TagName.Trim().ToUpper() == "IMG" && doc.All[i].Id != null && doc.All[i].Id.Trim().ToUpper() == "MANGAFILE")
                {
                    string Value = doc.All[i].GetAttribute("src");//取得SRC
                    workoption.WB.Stop();
                    //下載
                    string FE = ".jpg";
                    int Index01 = Value.LastIndexOf(".");
                    int Index02 = Value.IndexOf("?", Index01);
                    if (Index02 > -1)
                        FE = Value.Substring(Index01, Index02 - Index01);//副檔名
                    else
                        FE = Value.Substring(Index01);//副檔名
                    if (FE.Length > 4)
                        FE = ".jpg";
                    try
                    {
                        string SubDirPath = workoption.DirPath + workoption.MainName + workoption.Name;
                        string FilePath = SubDirPath + "\\" + workoption.PageIndex.ToString().PadLeft(3, '0') + FE;//組合檔名
                        if (!Directory.Exists(SubDirPath + "\\"))
                            Directory.CreateDirectory(SubDirPath + "\\");

                        //
                        string ErrorMessage = DoDownLoadPic(ref workoption, Value, FilePath, FE);


                        #region 有錯誤時
                        if (ErrorMessage.Trim() != "")//有錯誤時
                        {
                            if (workoption.RetryCnt <= 3)//重試
                            {
                                if (workoption.RetryCnt > 2)//已重試3次時重置執行緒狀態 重新下載
                                {
                                    workoption.NavigateCNT += 20;
                                }
                                workoption.RetryCnt++;
                                DoDownLoadWork(workoption);//再試一次
                                //DoDownLoadWorkobj(workoption);//繼續下載※
                                return;
                            }
                            else
                            {
                                workoption.WB.Stop();
                                workoption.WB.Dispose();
                                if (workoption.thread != null && workoption.thread.IsAlive)
                                {
                                    //將執行序停止並釋放
                                    workoption.thread.Abort();
                                    workoption.thread.Join();
                                }
                                workoption.Message = "錯誤，重試" + workoption.RetryCnt;
                                workoption.IsWork = 0;
                                workoption.Status = 4;
                                CheckWorkAndStartWork();//更新狀態(下載內頁圖片 錯誤，重試)
                                //執行下一個工作

                                return;
                            }


                        } 
                        #endregion
                        workoption.RetryCnt = 0;
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show("錯誤", ex.Message);
                        workoption.WB.Stop();
                        workoption.WB.Dispose();
                        //將執行序停止並釋放
                        if (workoption.thread != null && workoption.thread.IsAlive)
                        {
                            workoption.thread.Abort();
                            workoption.thread.Join();
                        }
                        workoption.IsWork = 0;//無執行
                        workoption.Status = 4;//錯誤停止
                        workoption.Message = "停止，錯誤05:" + ex.Message;
                        CheckWorkAndStartWork();//更新狀態(下載內頁圖片 Exception錯誤)
                        return;
                    }

                    //下一頁
                    workoption.PageIndex++;
                    if (workoption.PageCount == 0 || workoption.PageIndex > workoption.PageCount)
                    {
                        workoption.PageIndex = workoption.PageCount;
                        workoption.WB.Stop();
                        workoption.WB.Dispose();
                        //將執行序停止並釋放
                        if (workoption.thread != null && workoption.thread.IsAlive)
                        {
                            workoption.thread.Abort();
                            workoption.thread.Join();
                        }
                        workoption.IsWork = 0;//無執行
                        workoption.Status = 2;//已完成
                        workoption.Message = "完成下載";
                        if (!chkShowComplete.Checked)//如果不Show完成下載
                            workoption.FLP.Visible = false;
                        if (chkRecord.Checked)
                        {
                            string strNUM = "";
                            string strSUB_NUM = "";
                            string P = "";
                            string strURL = "";
                            GetNumURL(workoption.URL, out strURL, out strNUM, out strSUB_NUM, out P);
                            AddRecord(strURL, strNUM, workoption.Name, strSUB_NUM);
                        }
                        CheckWorkAndStartWork();//更新狀態(下載內頁圖片 完成下載)
                        return;
                    }

                    //ShowWorkList();//更新狀態
                    DoDownLoadWork(workoption);//繼續下載
                    //DoDownLoadWorkobj(workoption);//繼續下載※
                    return;
                }
            }
            //都沒找到IMG時
            if (workoption.RetryCnt <= 3)//重試
            {
                workoption.RetryCnt++;
                workoption.Message = "重試第" + workoption.PageIndex + "頁";
                DoDownLoadWork(workoption);//
                //DoDownLoadWorkobj(workoption);//繼續下載※
            }
            else
            {
                workoption.WB.Stop();
                workoption.WB.Dispose();
                //將執行序停止並釋放
                if (workoption.thread != null && workoption.thread.IsAlive)
                {
                    workoption.thread.Abort();
                    workoption.thread.Join();
                }
                workoption.IsWork = 0;//無執行
                workoption.Status = 4;//錯誤停止
                workoption.Message = "停止，錯誤06";
                CheckWorkAndStartWork();//更新狀態(下載內頁圖片 Exception錯誤)
                return;
            } 
            #endregion
        }


        //呼叫Process command
        public string RunProcess(string command)
        {
            string ALL = "";
            StringBuilder sb = new StringBuilder();
            string version = System.Environment.OSVersion.VersionString;//读取操作系统版本  
            if (version.Contains("Windows"))
            {
                using (Process p = new Process())
                {
                    p.StartInfo.FileName = "cmd.exe";

                    p.StartInfo.UseShellExecute = false;//是否指定操作系统外壳进程启动程序  
                    p.StartInfo.RedirectStandardInput = true;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.CreateNoWindow = true;//不显示dos命令行窗口  

                    p.Start();//启动cmd.exe  
                    p.StandardInput.WriteLine(command);//输入命令  
                    p.StandardInput.WriteLine("exit");//退出cmd.exe  
                    //p.WaitForExit();//等待执行完了，退出cmd.exe 

                    using (StreamReader reader = p.StandardOutput)//截取输出流  
                    {
                        //string line = reader.ReadLine();//每次读取一行  
                        ALL = reader.ReadToEnd();//每次读取一行  
                        //while (!reader.EndOfStream)
                        //{
                        //    sb.Append(line).Append("\r\n");//在Web中使用<br />换行  
                        //    line = reader.ReadLine();
                        //}
                        p.WaitForExit();//等待程序执行完退出进程  
                        p.Close();//关闭进程  
                        reader.Close();//关闭流  
                    }

                     
                }
            }
            //return sb.ToString();
            return ALL;
        }
        private string DoDownLoadPic_ByWebRequest(ref WorkOption workoption, string strURL, string FilePath, string FE)
        {
            string strErrorMessage = "";
            try
            {
                if (File.Exists(FilePath))
                {
                    switch (workoption.SaveProcessType)//內頁重覆的處理方式 0跳過 1覆蓋 2重新命名
                    {
                        case 0:
                            return "";
                        case 1:
                            File.Delete(FilePath);
                            break;
                        case 2:
                            FileReName(ref FilePath);
                            break;
                        default:
                            return "";
                    }
                }


                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strURL);
                request.Headers.Clear();
                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36";
                request.Referer = "https://www.manhuagui.com/";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream stream = response.GetResponseStream();
                byte[] Result = ReadStream(stream, 32765);
                stream.Close();
                FileStream myFile = File.Open(FilePath, FileMode.Create);
                myFile.Write(Result, 0, Result.Length);
                myFile.Dispose();
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message;
            }

            return strErrorMessage;
        }

        //下載圖片(截圖)
        private string DoDownLoadPic(ref WorkOption workoption,string strURL, string FilePath,string FE)
        {
            string ErrorMessage = "";

            workoption.WB.ScrollBarsEnabled = false;
            workoption.WB.Stop();

            ////ReNavigate(ref workoption);
            ////重新導頁
            //string strFinalURL = string.Format("{0}{1}{2}", workoption.URL.Trim(), "#p=", 1);
            //workoption.WB.Navigate(strFinalURL);//導頁
            //int tryCount2 = 0;
            //loading(workoption.WB, ref tryCount2);//等待完成載入

            IniWebBrowser();
            workoption.WB.Navigate(strURL);////導至圖片頁
            int tryCount = 0;
            loading(workoption.WB, ref tryCount);//等待完成載入
            IncludeScript(workoption.WB);//加入Script
            workoption.WB.Document.BackColor = Color.FromArgb(backColorR, backColorG, backColorB);
            try
            {
                if (File.Exists(FilePath))
                {
                    switch (workoption.SaveProcessType)//內頁重覆的處理方式 0跳過 1覆蓋 2重新命名
                    {
                        case 0:
                            return "";
                        case 1:
                            File.Delete(FilePath);
                            break;
                        case 2:
                            FileReName(ref FilePath);
                            break;
                        default:
                            return "";
                    }
                }

                #region 判斷檔案類型
                System.Drawing.Imaging.ImageFormat IF;
                switch (FE.ToUpper().Trim())
                {
                    case ".PNG":
                        IF = System.Drawing.Imaging.ImageFormat.Png;
                        break;
                    case ".BMP":
                        IF = System.Drawing.Imaging.ImageFormat.Bmp;
                        break;
                    case ".EMF":
                        IF = System.Drawing.Imaging.ImageFormat.Emf;
                        break;
                    case ".EXIF":
                        IF = System.Drawing.Imaging.ImageFormat.Exif;
                        break;
                    case ".GIF":
                        IF = System.Drawing.Imaging.ImageFormat.Gif;
                        break;
                    case ".ICON":
                        IF = System.Drawing.Imaging.ImageFormat.Icon;
                        break;
                    case ".JPEG":
                    case ".JPG":
                        IF = System.Drawing.Imaging.ImageFormat.Jpeg;
                        break;
                    case ".TIFF":
                        IF = System.Drawing.Imaging.ImageFormat.Tiff;
                        break;
                    case ".WMF":
                        IF = System.Drawing.Imaging.ImageFormat.Wmf;
                        break;
                    default:
                        IF = System.Drawing.Imaging.ImageFormat.Jpeg;
                        break;
                }
                #endregion

                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件
                sw.Reset();//碼表歸零
                sw.Start();//碼表開始計時
                int i = 0;
                while (sw.Elapsed.TotalMilliseconds < workoption.SleepMilliseconds)
                {
                    i++;
                }

                sw.Stop();//碼錶停止
                          //印出所花費的總豪秒數
                string result1 = sw.Elapsed.TotalMilliseconds.ToString();


                Bitmap bitmap = new Bitmap(workoption.Width, workoption.Height);
                Rectangle rectangle = new Rectangle(0, 0, workoption.Width, workoption.Height);  // 繪圖區域
                workoption.WB.DrawToBitmap(bitmap, rectangle);
                //裁切圖片
                bitmap = AutoChangeSize(bitmap);

                bitmap.Save(FilePath, IF);

                //另存圖片
                //workoption.WB.ShowSaveAsDialog();

                //System.IO.StreamReader reader = new System.IO.StreamReader(workoption.WB.DocumentStream, System.Text.Encoding.Default);
                //string gethtml = reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                //if (ErrorMessage.IndexOf("參數無效") >= 0 && RetryCnt <= 3)
                //{
                //    RetryCnt++;
                //    DoDownLoadPic(strURL, FilePath);
                //}
                //throw;
            }

            return ErrorMessage;
        }

        private Bitmap AutoChangeSize(Bitmap bitmap)
        {
            try
            {
                int Width = bitmap.Width;
                int Height = bitmap.Height;
                int Width_MAX = 0;
                int Height_MAX = 0;
                int Width_MIN = 99999;
                int Height_MIN = 99999;
                //Stopwatch sw = new Stopwatch();
                //sw.Reset();
                //sw.Start();
                //新方式 斜角掃瞄
                int tangleNum = Width >= Height ? Width : Height;

                int R = backColorR;
                int G = backColorG;
                int B = backColorB;
                //左上的點往斜下方向掃描
                int interval = 2;
                for (int t = 0; t < tangleNum; t += interval)
                {
                    //掃瞄寬
                    Color C_t = bitmap.GetPixel(t, t);//
                    if (C_t != null && (C_t.R != R || C_t.G != G || C_t.B != B))//有找到圖片的點
                    {
                        int point = t - interval + 1;
                        if (point < 0)
                            point = 0;
                        for (int n = point; n <= t; n++)//refind再找一次正確的左上啟始點
                        {
                            Color C_try = bitmap.GetPixel(n, n);//此第t個點
                            if (C_try != null && (C_try.R != R || C_try.G != G || C_try.B != B))//有第一個點時
                            {
                                //find min width
                                for (int m = 0; m <= n; m++)
                                {
                                    Color C_try2 = bitmap.GetPixel(m, n);//
                                    if (C_try2.R != R || C_try2.G != G || C_try2.B != B)
                                    {
                                        Width_MIN = m;
                                        break;
                                    }
                                }
                                //find max width
                                for (int m = tangleNum - 1; m >= n; m--)
                                {
                                    Color C_try2 = bitmap.GetPixel(m, n);//
                                    if (C_try2 != null && (C_try2.R != R || C_try2.G != G || C_try2.B != B))
                                    {
                                        Width_MAX = m;
                                        break;
                                    }
                                }
                                //find min height
                                for (int m = 0; m <= n; m++)
                                {
                                    Color C_try2 = bitmap.GetPixel(n, m);//
                                    if (C_try2.R != R || C_try2.G != G || C_try2.B != B)
                                    {
                                        Height_MIN = m;
                                        break;
                                    }
                                }
                                //find max height
                                for (int m = tangleNum - 1; m >= n; m--)
                                {
                                    Color C_try2 = bitmap.GetPixel(n, m);//
                                    if (C_try2 != null && (C_try2.R != R || C_try2.G != G || C_try2.B != B))
                                    {
                                        Height_MAX = m;
                                        break;
                                    }
                                }
                                goto End;
                            }
                        }
                    }
                }

                End:

                int newWidth = Width_MAX - Width_MIN + 1;
                int newHeight = Height_MAX - Height_MIN + 1;
                if (newWidth <= 0)
                    newWidth = 1;
                if (newHeight <= 0)
                    newHeight = 1;
                Bitmap bmp = new Bitmap(newWidth, newHeight, PixelFormat.Format24bppRgb);
                Graphics g = Graphics.FromImage(bmp);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                //g.DrawImage(bitmap, new Rectangle(0, 0, Width_MAX, Height_MAX), new Rectangle(0, 0, Width_MAX, Height_MAX), GraphicsUnit.Pixel);
                //g.DrawImage(bitmap,0,0, new Rectangle(0, 0, Width_MAX, Height_MAX), GraphicsUnit.Pixel);
                g.DrawImage(bitmap, 0, 0, new Rectangle(Width_MIN, Height_MIN, newWidth, newHeight), GraphicsUnit.Pixel);

                g.Dispose();
                return bmp;
            }
            catch (Exception ex)
            {

                throw;
            }
        }


        private void FileReName(ref string FilePath)
        {
            if (File.Exists(FilePath))
            {
                int IndexS = FilePath.LastIndexOf("\\")+1;
                int IndexE = FilePath.LastIndexOf(".");
                string FileName = FilePath.Substring(IndexS, IndexE - IndexS);
                FileName += "_";
                FilePath = FilePath.Substring(0, IndexS) + FileName + FilePath.Substring(IndexE);
                FileReName(ref FilePath);
            }
        }

        public void loading(WebBrowser WB,ref int tryCount)
        {
            DateTime dateNow = DateTime.Now;
            try
            {
             
                while (!(WB.ReadyState == WebBrowserReadyState.Complete) && (DateTime.Now - dateNow).TotalSeconds <= 20)
                {
                    Thread.Sleep(200);
                    Application.DoEvents();
                    Thread.Sleep(200);
                }
         

            }
            catch (Exception ex)
            {
                tryCount++;
                if (tryCount < 3)
                {
                    loading(WB, ref tryCount);
                    return;
                }

                WB.Stop();
                WB.Dispose();
                MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                return;
            }
        }

        private void btnAllSEL_Click(object sender, EventArgs e)//全選
        {
            DataTable dt = (DataTable)gvList.DataSource;
            dt.AsEnumerable().ToList().ForEach(X => X["SEL"] = "True");
            gvList.DataSource = dt;
        }
        private void btnUnSEL_Click(object sender, EventArgs e)//反選
        {
            DataTable dt = (DataTable)gvList.DataSource;
            dt.AsEnumerable().ToList().ForEach(X=>X["SEL"] = X["SEL"].ToString().Trim().ToUpper() == "TRUE" ? "False" : "True");
            gvList.DataSource = dt;
        }

        private void btnBrowser_Click(object sender, EventArgs e)//選擇預設儲存路徑
        {

            DialogResult dialogresult = folderBrowserDialog2.ShowDialog();
            if (dialogresult == DialogResult.OK)
                txtDirPath.Text = folderBrowserDialog2.SelectedPath;
        }

        private void btnSaveSaveSetting_Click(object sender, EventArgs e)//儲存預設儲存路徑設定
        {
            try
            {
                GetSettingItem g = new GetSettingItem();
                if (chkIsUseDefaultSavePath.Checked)
                {
                    g.SetValue("IsUseDefaultSavePath", "1");
                    g.SetValue("DefaultSavePath", txtDirPath.Text.Trim());
                }
                else
                    g.SetValue("IsUseDefaultSavePath", "0");

                MessageBox.Show("儲存成功!", "完成", MessageBoxButtons.OK, MessageBoxIcon.None);//狀態列
            }
            catch (Exception ex)
            {

                MessageBox.Show("儲存失敗!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
            }

        }

        private void btnClearWorkList_Click(object sender, EventArgs e)//清除已完成工作清單
        {
            List<WorkOption> listWorkOptionTemp = new List<WorkOption>(listWorkOption);
            //只有已完成的可以清除
            var v = from WorkOption t in listWorkOption
                    where t.Status == 2
                    select t;
            foreach (var V in v)
            {
                if (V.WB != null)
                {
                    V.WB.Stop();
                    V.WB.Dispose();
                }
                if (V.thread != null && V.thread.IsAlive)
                {
                    try
                    {
                        //將執行序停止並釋放
                        V.thread.Abort();
                        V.thread.Join();
                    }
                    catch (Exception)
                    {

                    }
                }
                //刪除項目
                listWorkOptionTemp.Remove(V);

            }
            listWorkOption = listWorkOptionTemp;

            FLPWorkList.Controls.Clear();
            ShowWorkList();//顯示項目


        }

        private void btnGoOnAll_Click(object sender, EventArgs e)//全部繼續
        {
            var v = from WorkOption t in listWorkOption
                    where t.Status == 3
                    select t;
            foreach (WorkOption workoption in v)
            {
                //暫停中 換成 進行中
                workoption.IsWork = 0;
                workoption.Status = 0;

            }

            CheckWorkAndStartWork();//(全部繼續)
        }

        private void btnStopAll_Click(object sender, EventArgs e)//全部暫停
        {

            var v = from WorkOption t in listWorkOption
                    where t.Status != 2 && t.Status != 3
                    select t;//非已完成、已暫停
            if (v == null || v.Count() < 1) return;
            foreach (WorkOption workoption in v)
            {
                if (workoption == null) continue;

                if (workoption.Status == 1)//進行中 換成 暫停
                {
                    workoption.IsWork = 2;//暫停
                    workoption.Status = 3;//暫停
                    if (workoption.thread != null && workoption.thread.IsAlive)
                        workoption.thread.Interrupt();

                }
                else if (workoption.Status == 0)//未執行 換成 暫停
                {
                    workoption.IsWork = 2;//暫停
                    workoption.Status = 3;//暫停
                }

            }

            CheckWorkAndStartWork();//(全部暫停)
        }

        private void btnAllDelete_Click(object sender, EventArgs e)//全部刪除
        {
            try
            {
                DialogResult result = MessageBox.Show("確定刪除所有工作?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                if (result == System.Windows.Forms.DialogResult.Cancel)//取消
                    return;

                //是否還有工作執行中
                var v = from WorkOption t in listWorkOption
                        where t.Status == 1
                        select t;
                if (v != null && v.Count() > 0)
                {
                    DialogResult result2 = MessageBox.Show("工作在進行中確定要刪除?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result2 == System.Windows.Forms.DialogResult.Cancel)//取消
                        return;
                }
                if (listWorkOption != null && listWorkOption.Count() > 0)
                {
                    foreach (WorkOption workoption in listWorkOption)
                    {
                        int Status = workoption.Status;
                        if (Status == 1)//進行中
                        {
                            try
                            {
                                if (workoption.thread != null && workoption.thread.IsAlive)
                                {
                                    if (workoption.WB != null)
                                    {
                                        workoption.WB.Stop();
                                        workoption.WB.Dispose();
                                    }
                                    workoption.thread.Interrupt();
                                    workoption.thread.Abort();
                                    workoption.thread.Join();//用來等候Thread投擲出"使Thread停止的exception
                                    listWorkOption.Remove(workoption);//從清單內刪除
                                }
                                Thread.Sleep(100);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("工作執行中，錯誤無法刪除!", "警告", MessageBoxButtons.OK, MessageBoxIcon.None);
                            }
                        }
                        //else if (Sataus == 2)//已完成
                        //{
                        //    MessageBox.Show("工作已完成無法刪除!", "警告", MessageBoxButtons.OK, MessageBoxIcon.None);
                        //    return;
                        //}
                        else//其他狀態 刪除
                            listWorkOption.Remove(workoption);//從清單內刪除
                    }
                    
                }
                listWorkOption.Clear();

                FLPWorkList.Controls.Clear();
                MessageBox.Show("刪除工作成功!", "成功", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
            catch (Exception ex)
            {
                if (ex.Message.IndexOf("集合") > -1)
                {
                    listWorkOption.Clear();
                    ShowWorkList();//刪除工作後顯示新清單
                    return;
                }
                MessageBox.Show(ex.Message);
            }

        }

        #region 内存回收
        [DllImport("kernel32.dll", EntryPoint = "SetProcessWorkingSetSize")]

        public static extern int SetProcessWorkingSetSize(IntPtr process, int minSize, int maxSize);
        /// <summary>
        /// 释放内存
        /// </summary>
        public static void ClearMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
            }

        }

        #endregion


        #region 縮至工具列
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (WorkOption workoption in listWorkOption)
            {
                if (workoption.WB != null)
                {
                    workoption.WB.Stop();
                    workoption.WB.Dispose();
                }
                if (workoption.thread != null && workoption.thread.IsAlive)
                {
                    try
                    {
                        //將執行序停止並釋放
                        workoption.thread.Abort();
                        workoption.thread.Join();
                    }
                    catch (Exception)
                    {

                    }
                }

                ClearMemory();//記憶體回收
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)//縮小時
            {
                this.notifyIcon1.Visible = true;//顯示Icon
                this.Hide();//隱藏Form
                //if (ReadFileTimer.Enabled)//有執行時 眨眼
                //{
                //    Eye_timer.Enabled = true;//開始眨眼
                //    Eye_timer.Start();//開始眨眼
                //}
            }
            else//放大時
            {
                this.notifyIcon1.Visible = false;//隱藏Icon
                //Eye_timer.Stop();//停止眨眼
                //Eye_timer.Enabled = false;//停止眨眼
            }

            
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();//顯示Form
            this.WindowState = FormWindowState.Normal;//回到正常大小
            this.Activate();//焦點
            this.Focus();//焦點
        }

        private void notifyIcon1_MouseMove(object sender, MouseEventArgs e)
        {
            var vWorkOption = from WorkOption t in listWorkOption
                              where t.IsWork == 1
                              select t;
            
            string Info = labStatus.Text;
            string[] Infos = Info.Split(new string[] { "，" }, StringSplitOptions.RemoveEmptyEntries);
            if (Infos.Length > 3)
            {
                string NewInfo = "";
                for (int i = 0; i < Infos.Length; i++)
                {
                    if (i == 0)
                    {
                        string strKey = "個工作";
                        int Index01 = Infos[i].IndexOf(strKey);
                        if (Index01 > -1)
                            NewInfo += Infos[i].Substring(Index01 + strKey.Length)+"\r\n";
                        else
                            NewInfo += Infos[i] + "\r\n";
                    }
                    else
                    {
                        NewInfo += Infos[i].Replace("個工作", "") + "\r\n";
                    }
                }
                Info = NewInfo;
            }
            string Other = "";
            if (vWorkOption != null && vWorkOption.Count() > 0)
            {
                WorkOption workoption = vWorkOption.ToList()[0];
                string MainName = workoption.MainName.Replace("\\", "");
                if (MainName.Length > 5)
                    MainName = MainName.Substring(0, 5) + "...";
                Other = string.Format("{0} {1}\r\n{2}/{3}", MainName, workoption.Name, workoption.PageIndex, workoption.PageCount);
            }
             notifyIcon1.Text = Info+Other;
        }
        #endregion

        //顯示已完成工作
        private void chkShowComplete_CheckedChanged(object sender, EventArgs e)
        {
            if(chkShowComplete.Checked)//顯示已完成工作
            {
                FLPWorkList.Controls.Clear();
                ShowWorkList(false);//顯示項目
            }
            else//不顯示已完成工作
            {
                var v2 = from WorkOption t in listWorkOption
                         where t.Status != 2
                         select t;

                FLPWorkList.Controls.Clear();
                ShowWorkList(false);//顯示項目
            }
        }


        #region 導頁靜音
        const int FEATURE_DISABLE_NAVIGATION_SOUNDS = 21;
        const int SET_FEATURE_ON_PROCESS = 0x00000002;

        [DllImport("urlmon.dll")]
        [PreserveSig]
        [return: MarshalAs(UnmanagedType.Error)]
        static extern int CoInternetSetFeatureEnabled(
            int FeatureEntry,
            [MarshalAs(UnmanagedType.U4)] int dwFlags,
            bool fEnable);

        static void DisableClickSounds()
        {
            CoInternetSetFeatureEnabled(
                FEATURE_DISABLE_NAVIGATION_SOUNDS,
                SET_FEATURE_ON_PROCESS,
                true);
        }

        #endregion

        static DataTable dtTemp = new DataTable();
        private void btnSearch_Click(object sender, EventArgs e)//搜尋關鍵字
        {
            DataTable dt = dtTemp;
            string keyWord = txtSearch.Text.Trim();
            var vdt = from t in dt.AsEnumerable()
                      where t["Name"].ToString().Trim().IndexOf(keyWord) > -1
                      select t;
            if (vdt != null && vdt.Count() > 0)
                dt = vdt.CopyToDataTable();
            else
                dt = dt.Clone();
            gvList.DataSource = dt;
        }

        private void FLPWorkList_MouseEnter(object sender, EventArgs e)
        {
            FLPWorkList.Focus();
        }

        private void gvList_MouseEnter(object sender, EventArgs e)
        {
            gvList.Focus();
        }
        float FormWidthLog;
        private void btnOpenClosePanel_Click(object sender, EventArgs e)
        {
            if (btnOpenClosePanel.ImageIndex == 0)//打開時 關閉
            {
                FormWidthLog = this.Width - tlpBase.ColumnStyles[0].Width;
                this.Width = (int)tlpBase.ColumnStyles[0].Width+20;
                tlpBase.ColumnStyles[1].SizeType = SizeType.Absolute;
                tlpBase.ColumnStyles[1].Width = 0;

                tlpBase.ColumnStyles[0].SizeType = SizeType.Percent;
                tlpBase.ColumnStyles[0].Width = 100;
                btnOpenClosePanel.ImageIndex = 1;
            }
            else//關閉時 打開
            {
                if(wbSearch.Url==null)
                    wbSearch.Navigate("https://www.manhuagui.com/");//https://tw.manhuagui.com/
                this.Width = (int)(790 + FormWidthLog);
                tlpBase.ColumnStyles[1].Width = 100;
                tlpBase.ColumnStyles[1].SizeType = SizeType.Percent;

                tlpBase.ColumnStyles[0].SizeType = SizeType.Absolute;
                tlpBase.ColumnStyles[0].Width = 790;
                btnOpenClosePanel.ImageIndex = 0;
            }
        }

        private void btnAddWork_Click(object sender, EventArgs e)
        {
            string strURL = wbSearch.Url.ToString();
            if (!CheckURL(strURL)) return;
            txtURL.Text = strURL;
            btnSetting.PerformClick();

        }

        //加入最愛按鈕
        private void btnAddOphen_Click(object sender, EventArgs e)
        {
            string strURL = wbSearch.Url.ToString();
            if (!CheckURL(strURL)) return;
            Thread thread = new Thread(AddOphen);
            thread.IsBackground = true;//設為背景執行~才可刪除
            thread.Start(new object[] { strURL, lvOphen, imlItem });//執行 執行緒

        }

        private delegate void delegateAddOphen(object obj);
        private void AddOphen(object obj)//增加最愛項目
        {
            if (this.InvokeRequired)//對跨執行緒做處理
            {
                delegateAddOphen myUpdate = new delegateAddOphen(AddOphen);
                this.Invoke(myUpdate, obj);
            }
            else
            {
                string strURL = ((object[])obj)[0].ToString();//傳入的參數
                ListView lvOphen_ = (ListView)((object[])obj)[1];//傳入的參數
                ImageList imlItem_ = (ImageList)((object[])obj)[2];//傳入的參數
                AddOphen(strURL, ref lvOphen_, ref imlItem_);
            }
        }

        //增加最愛項目
        private void AddOphen(string strURL, ref ListView lvOphen_, ref ImageList imlItem_)
        {

            try
            {
                if (lvOphen_ == null)
                    lvOphen_ = this.lvOphen;
                if (imlItem_ == null)
                    imlItem_ = this.imlItem;
                string strNUM = "";
                string strSUB_NUM = "";
                string P = "";
                GetNumURL(strURL, out strURL, out strNUM, out strSUB_NUM, out P);

                List<ListViewItem> lvItemItems = new List<ListViewItem>();//有連結存在的項目
                //比對並新增縮圖項目
                GetURLIcon(strURL, strNUM, strSUB_NUM, ref lvOphen_, ref imlItem, ref lvItemItems);
                                                                                             



                ListViewToCSV(lvOphen_, OphenFileName);

            }
            catch (Exception ex)
            {


            }

        }

        //比對並新增縮圖項目
        public void GetURLIcon(string strURL, string strNUM, string strSUB_NUM, ref ListView lvItem, ref ImageList imlItem, ref List<ListViewItem> lvItemItems)
        {
            var vlvItem = from ListViewItem t in lvItem.Items
                          where t.Tag.ToString().Trim() == strURL.Trim()
                          select t;
            if (vlvItem != null && vlvItem.Count() > 0)//若已經存在就跳過
            {
                lvItemItems.AddRange(vlvItem);
                return;

            }

            //將ListVlew Item塞資料
            ListViewItem Item = new ListViewItem();
            GetBookInfoList(strURL, strNUM, ref Item, ref imlItem);


            if (this.lvOphen.InvokeRequired)//在非當前執行緒內 使用委派
            {
                UpdatelvItemProcess myUpdate = new UpdatelvItemProcess(SetlvItem);
                lvOphen.Invoke(myUpdate, Item);
            }
            else
            {
                lvOphen.Items.Add(Item);
            }

            lvItemItems.Add(Item);
            Thread.Sleep(100);





        }

        public void GetBookInfoList(string strURL, string strNUM, ref ListViewItem Item, ref ImageList imlItem)
        {
            DataTable dtBOOK = new DataTable();//所有卷(話)
            GetBookInfoList(strURL, strNUM, ref Item, ref imlItem, out dtBOOK,false, false);
        }
        //ListView取得連結標題跟圖片URL
        public void GetBookInfoList(string strURL, string strNUM, ref ListViewItem Item,ref ImageList imlItem,out DataTable dtBOOK,bool IsHaveBookRecord,bool IsGetBook)
        {
            dtBOOK = new DataTable();//所有卷(話)
         
            WebClient wc = new WebClient();
            wc.Encoding = Encoding.UTF8;
            string Contents = "";
            try
            {
                Contents = wc.DownloadString(strURL);
            }
            catch (Exception ex)
            {
                //BookInfo.WORK_STATUS = 3;
                //BookInfo.WORK_STATUS_MESSAGE = string.Format("編號{0}錯誤[{1}]", NUM.ToString(), ex.Message);
                return;
            }

            string NewContents = "";
            string PICTURE = "", NAME = "", YEAR = "", ZONE = "", LETTER = "", TYPE_SUM = "", WRITER = "",
                NAME2 = "", STATUS = "", UPDATE_TIME = "", NEW_PAGE = "", DOCUMENT = "";



            #region 取得各項資訊
            //更新集數資料
            try
            {
                if (IsGetBook)
                    //所有集數
                    dtBOOK = GetAllBookDataTable(strURL, Contents);
            }
            catch (Exception ex)
            { }
            if (IsHaveBookRecord)//已經建過機本資料時跳過
                return;
            try
            {
                //圖片
                PICTURE = GetPICTURE(Contents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //取得名稱
                NAME = GetTitle(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //出品年代
                YEAR = GetYEAR(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //漫畫地區
                ZONE = GetZONE(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //字母索引
                LETTER = GetLETTER(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //漫畫劇情 多個
                TYPE_SUM = GetTYPE_SUM(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //漫畫作者
                WRITER = GetWRITER(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //漫畫別名
                NAME2 = GetNAME2(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //漫畫狀態
                STATUS = GetSTATUS(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //更新時間
                UPDATE_TIME = GetUPDATE_TIME(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //最新一集
                NEW_PAGE = GetNEW_PAGE(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
            try
            {
                //介紹
                DOCUMENT = GetDOCUMENT(NewContents.Trim() == "" ? Contents : NewContents, out NewContents);
            }
            catch (Exception ex)
            { }
           
            #endregion

            //NAME = _NAME;//名稱
            //YEAR = _YEAR;//出品年代
            //ZONE = _ZONE;//漫畫地區
            //LETTER = _LETTER;//字母索引
            //TYPE_SUM = _TYPE_SUM;//漫畫劇情 多個
            //WRITER = _WRITER;//漫畫作者
            //NAME2 = _NAME2;//漫畫別名
            //STATUS = _STATUS;//漫畫狀態
            //UPDATE_TIME = _UPDATE_TIME;//更新時間
            //NEW_PAGE = _NEW_PAGE;//最新一集
            //DOCUMENT = _DOCUMENT;//介紹
            //PICTURE = _PICTURE;//圖片
            //listBOOK = _listBOOK;//目前所有卷/回 數

            string picPath = "";
            if (PICTURE.Trim() != "")
            {
                //圖片檔
                string picDirPath = AppDomain.CurrentDomain.BaseDirectory + "\\Book_Picture\\";
                if (!Directory.Exists(picDirPath))
                    Directory.CreateDirectory(picDirPath);
                picPath = picDirPath + strNUM.ToString() + ".jpg";
                if (!File.Exists(picPath))
                {
                    wc.DownloadFile(PICTURE, picPath);
                }
            }
            

            //ImageList增加項目
            try
            {
                imlItem.Images.Add(new Bitmap(picPath));
            }
            catch (Exception ex)
            {

            }

            //ListView增加項目
            Item = new ListViewItem(new string[] { NAME , YEAR , ZONE , LETTER , TYPE_SUM , 
                                                   WRITER ,NAME2 , STATUS , UPDATE_TIME , NEW_PAGE , 
                                                   DOCUMENT,PICTURE,strURL,strNUM }, imlItem.Images.Count - 1);
            Item.Text = NAME;
            Item.ToolTipText = strURL;
            Item.Tag = strURL;



        }

        //移除lvItem多餘的項目
        public void CheckClearListView(ref ListView lvItem, List<ListViewItem> lvItemItems)
        {
            //移除lvItem多餘的項目
            if (lvItemItems != null && lvItemItems.Count > 0)
            {
                foreach (ListViewItem item in lvItem.Items)
                {
                    var v = from ListViewItem t in lvItemItems
                            where t.ToolTipText.Trim().Contains(item.ToolTipText.Trim())
                            select t;
                    if (v == null || v.Count() < 1)
                        lvItem.Items.Remove(item);
                }

            }
            else
                lvItem.Items.Clear();
        }

        //從ListView存成CSV檔
        public void ListViewToCSV(ListView lvOphen_,string SettingFileName)
        {

            try
            {
                string strCSV = "NO,NAME,YEAR,ZONE,LETTER,TYPE_SUM,WRITER,NAME2,STATUS,UPDATE_TIME,NEW_PAGE,DOCUMENT,PICTURE,URL,NUM";
                var v = from ListViewItem t in lvOphen_.Items
                        select t;
                int NO = 1;
                foreach (ListViewItem item in v)
                {
                    strCSV += string.Format("\r\n{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14}",
                        NO.ToString(),
                        item.SubItems[0].Text,
                        item.SubItems[1].Text,
                        item.SubItems[2].Text,
                        item.SubItems[3].Text,
                        item.SubItems[4].Text,
                        item.SubItems[5].Text,
                        item.SubItems[6].Text,
                        item.SubItems[7].Text,
                        item.SubItems[8].Text,
                        item.SubItems[9].Text,
                        item.SubItems[10].Text,
                        item.SubItems[11].Text,
                        item.SubItems[12].Text,
                        item.SubItems[13].Text
                        );

                    NO++;

                }

                GetSettingItem getSettingItem = new GetSettingItem();
                getSettingItem.SetCSVString(strCSV, SettingFileName);

            }
            catch (Exception ex)
            {

                
            }
        }

        //從ListView存成CSV檔
        public void lvRecordBookToCSV(ListView lvOphen_, string strDirPath,string FileName)
        {

            try
            {
                string strCSV = "Name,PageCount,URL,SubNum,IsDownLoad,Status,DateTime";
                var v = from ListViewItem t in lvOphen_.Items
                        select t;
                int NO = 1;
                foreach (ListViewItem item in v)
                {
                    strCSV += string.Format("\r\n{0},{1},{2},{3},{4},{5},{6}",
                        item.SubItems[0].Text,
                        item.SubItems[1].Text,
                        item.SubItems[2].Text,
                        item.SubItems[3].Text,
                        item.SubItems[4].Text,
                        item.SubItems[5].Text,
                        item.SubItems[6].Text
                        );

                    NO++;

                }

                GetSettingItem getSettingItem = new GetSettingItem();
                getSettingItem.SetCSVString(strCSV,strDirPath,FileName);

            }
            catch (Exception ex)
            {


            }
        }

        public delegate void UpdatelvItemProcess(object method, object method2);
        public void SetlvItem(object Value, object Value2)
        {
            ListView listView = (ListView)Value;
            listView.Items.Add((ListViewItem)Value2);

        }
        //從CSV轉成最愛項目
        public void CSVToListView(ref ListView listView,ref ImageList imageList,string SettingFileName)
        {
            string Path = System.AppDomain.CurrentDomain.BaseDirectory + SettingFileName;

            if (!File.Exists(Path)) return;
            GetSettingItem getSettingItem = new GetSettingItem();
            DataTable dt = getSettingItem.GetCSVTable("", SettingFileName);

            listView.Items.Clear();
            imageList.Images.Clear();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                //ImageList增加項目
                try
                {
                    //圖片檔
                    string picDirPath = AppDomain.CurrentDomain.BaseDirectory + "\\Book_Picture\\";
                    if (!Directory.Exists(picDirPath))
                        Directory.CreateDirectory(picDirPath);
                    string picPath = picDirPath + dr["NUM"].ToString() + ".jpg";

                    imageList.Images.Add(new Bitmap(picPath));
                }
                catch (Exception ex)
                {

                }

                //ListView增加項目
                ListViewItem Item = new ListViewItem(new string[] {
                    dr["NAME"].ToString().Trim() ,
                    dr["YEAR"].ToString().Trim(),
                    dr["ZONE"].ToString().Trim(),
                    dr["LETTER"].ToString().Trim(),
                    dr["TYPE_SUM"].ToString().Trim(),
                    dr["WRITER"].ToString().Trim(),
                    dr["NAME2"].ToString().Trim(),
                    dr["STATUS"].ToString().Trim(),
                    dr["UPDATE_TIME"].ToString().Trim(),
                    dr["NEW_PAGE"].ToString().Trim(),
                    dr["DOCUMENT"].ToString().Trim(),
                    dr["PICTURE"].ToString().Trim(),
                    dr["URL"].ToString().Trim(),
                    dr["NUM"].ToString().Trim()}, imageList.Images.Count - 1);
                Item.Text = dr["NAME"].ToString().Trim();
                Item.ToolTipText = dr["URL"].ToString().Trim();
                string strURL = dr["URL"].ToString().Trim();
                strURL += strURL.LastIndexOf("/") == strURL.Length-1 ? "" : "/";
                Item.Tag = strURL.Trim();



                if (listView.InvokeRequired)//在非當前執行緒內 使用委派
                {
                    UpdatelvItemProcess myUpdate = new UpdatelvItemProcess(SetlvItem);
                    listView.Invoke(myUpdate, listView, Item);
                }
                else
                {
                    listView.Items.Add(Item);
                }


                Thread.Sleep(100);
            }



        }
        //瀏覽網頁歷程
        List<string> listURLHistory = new List<string>();
        int curIndex = 0;//目前瀏覽網頁歷程的Index
        bool IsButtonEven = false;//是否為瀏覽功能按鈕事件
        string CurrentURL = "";//目前連結
        //上一頁
        private void btnBack_Click(object sender, EventArgs e)
        {
            wbSearch.Stop();
            //wbSearch.GoBack();
            IsButtonEven = true;
            gotoIndexUrl(-1);
        }
        //下一頁
        private void btnNext_Click(object sender, EventArgs e)
        {
            wbSearch.Stop();

            //wbSearch.GoForward();
            IsButtonEven = true;
            gotoIndexUrl(1);
        }
        //重整
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            wbSearch.Stop();
            //IsButtonEven = true;
            wbSearch.Refresh();//重新整理
        }
        //首頁
        private void btnGoHome_Click(object sender, EventArgs e)
        {
            wbSearch.Stop();
            wbSearch.Navigate("https://www.manhuagui.com/");

        }
        //導至
        private void btnNavigated_Click(object sender, EventArgs e)
        {
            if (txtAdderss.Text.Trim() == "") return;
            try
            {
                wbSearch.Stop();
                wbSearch.Navigate(txtAdderss.Text.Trim());
            }
            catch (Exception ex)
            {

                wbSearch.Stop();
            }
        }
        //增加歷程
        private void addUrl(string curUrl)
        {
            //Add a new Url
            if (listURLHistory.Count > 0 && listURLHistory.Count - 1 > curIndex)
                listURLHistory.RemoveRange(curIndex, listURLHistory.Count - curIndex - 1);
            if (listURLHistory.Count>0 && curUrl == listURLHistory[listURLHistory.Count - 1]) return;

            listURLHistory.Add(curUrl);
            curIndex = listURLHistory.Count - 1;

            //gotoUrl(curUrl);
        }
        //前往網頁
        private void gotoIndexUrl(int addindex)
        {
            if (curIndex+addindex > listURLHistory.Count - 1) return;
            if (curIndex + addindex < 0) return;
            curIndex += addindex;
            string curUrl = listURLHistory[curIndex];
            //wbSearch.Navigated -= wbSearch_Navigated;
            wbSearch.Navigate(curUrl);
            int tryCount = 0;
            loading3(wbSearch, ref tryCount);//等待完成載入


        }
        //導頁事件
        private void wbSearch_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            if (CurrentURL == wbSearch.Url.ToString())
                return;
            else
                CurrentURL = wbSearch.Url.ToString();
            if (!IsButtonEven)
                addUrl(wbSearch.Url.ToString());
            IsButtonEven = false;
        
        }
        //網頁瀏覽器Loading
        public void loading3(WebBrowser WB, ref int tryCount)
        {
            DateTime dateNow = DateTime.Now;
            try
            {

                //btnBack.Enabled = false;
                //btnNext.Enabled = false;
                //btnRefresh.Enabled = false;
                //btnGoHome.Enabled = false;
                while (!(WB.ReadyState == WebBrowserReadyState.Complete) && (DateTime.Now - dateNow).TotalSeconds <= 10)
                {
                    Application.DoEvents();
                }
                btnBack.Enabled = true;
                btnNext.Enabled = true;
                btnRefresh.Enabled = true;
                btnGoHome.Enabled = true;
            }
            catch (Exception ex)
            {
                tryCount++;
                if (tryCount < 3)
                {
                    loading(WB, ref tryCount);
                    return;
                }

                WB.Stop();
                WB.Dispose();
                MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);//狀態列
                return;
            }
        }
        //導頁完成時 顯示網址
        private void wbSearch_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            try
            {
                txtAdderss.Text = wbSearch.Url.ToString().Trim();
            }
            catch (Exception ex)
            {

            }
        }

        private void wB_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

            //HtmlElement head = ((WebBrowser)sender).Document.GetElementsByTagName("head")[0];
            //HtmlElement scriptEl = ((WebBrowser)sender).Document.CreateElement("script");
            //IHTMLScriptElement element = (IHTMLScriptElement)scriptEl.DomElement;
            //element.text = "function sayHello() { ";
            //element.text += System.IO.File.ReadAllText(Environment.CurrentDirectory + "\\js\\polyfills.js");
            //element.text += System.IO.File.ReadAllText(Environment.CurrentDirectory + "\\js\\webp-hero.bundle.js");
            //element.text += "var webpMachine = new webpHero.WebpMachine();webpMachine.polyfillDocument(); }";
            //head.AppendChild(scriptEl);
            //((WebBrowser)sender).Document.InvokeScript("sayHello");

        }

        #region 取得漫畫介紹頁所有資訊

        //取得該ul內所有資訊
        private void GetliInfo(string ul, ref DataTable dt, string strURL)
        {
            if (ul.IndexOf("<a href=\"") < 0 || ul.IndexOf("<span>") < 0 || ul.IndexOf("<i>") < 0) return;

            string URL;
            //取得連結
            string NewContents;//取得連結後的內容
            GetValue(ul, out NewContents, out URL, "<a href=\"", "\"");
            //取得卷名
            string Name;//卷名
            GetValue(NewContents, out Name, "<span>", "<i>");
            //取得頁數
            string PageCount;
            string NewContents2;//取得頁數後的內容
            GetValue(NewContents, out NewContents2, out PageCount, "<i>", "</i>");

            string OrgURL = strURL.Trim();
            string[] OrgURLs = OrgURL.Split(new string[] { "/" }, StringSplitOptions.None);
            string[] URLs = URL.Split(new string[] { "/" }, StringSplitOptions.None);
            string CombinURL = "";
            for (int i = 0; i < URLs.Length; i++)
            {
                if (URLs[i].Trim() == "") continue;
                int IndexFlg = OrgURLs.ToList().FindIndex(U => U.Trim() == URLs[i].Trim() && U.Trim() != "");
                if (IndexFlg >= 0)
                {
                    CombinURL = string.Join("/", OrgURLs, 0, IndexFlg) + "/" + string.Join("/", URLs, i, URLs.Length - i);
                    break;
                }

            }
            string SubNum = URLs[URLs.Length - 1].Substring(0, URLs[URLs.Length - 1].IndexOf("."));

            DataRow dr = dt.NewRow();
            dr["SEL"] = true;
            dr["Name"] = Name;
            dr["PageCount"] = PageCount.ToUpper().Trim().Replace("P", "");
            dr["URL"] = CombinURL;
            dr["SubNum"] = SubNum;
            dr["IsDownLoad"] = 0;
            dr["Status"] = "未下載";
            dt.Rows.Add(dr);

            if (NewContents2.IndexOf("<a href=\"") >= 0)
                GetliInfo(NewContents2, ref dt, strURL);

        }

        //取得該ul內所有資訊
        private void GetliInfo2(string ul, ref DataTable dt, string strURL)
        {
            if (ul.ToUpper().IndexOf("<A ") < 0 || ul.ToUpper().IndexOf("HREF") < 0 || ul.ToUpper().IndexOf("<SPAN>") < 0 || ul.ToUpper().IndexOf("<I>") < 0) return;
            string NewContents;//取得連結後的內容

            string Name;//卷名
            GetValue(ul, out NewContents, out Name, "title=\"", "\"");
            if (Name.Trim() == "")
                GetValue(ul, out NewContents, out Name, "title=", " ");

            string URL;//取得連結
            GetValue(NewContents, out NewContents, out URL, "href=\"", "\"");
            //取得卷名
            if (Name.Trim() == "")
                GetValue(NewContents, out Name, "<span>", "<i>");
            //取得頁數
            string PageCount;
            string NewContents2;//取得頁數後的內容
            GetValue(NewContents, out NewContents2, out PageCount, "<i>", "</i>");

            string OrgURL = strURL.Trim();
            string[] OrgURLs = OrgURL.Split(new string[] { "/" }, StringSplitOptions.None);
            string[] URLs = URL.Split(new string[] { "/" }, StringSplitOptions.None);
            string CombinURL = "";
            for (int i = 0; i < URLs.Length; i++)
            {
                if (URLs[i].Trim() == "") continue;
                int IndexFlg = OrgURLs.ToList().FindIndex(U => U.Trim() == URLs[i].Trim() && U.Trim() != "");
                if (IndexFlg >= 0)
                {
                    CombinURL = string.Join("/", OrgURLs, 0, IndexFlg) + "/" + string.Join("/", URLs, i, URLs.Length - i);
                    break;
                }

            }
            if (Name.Trim() != "" && PageCount.Trim() != "" && CombinURL.Trim() != "")
            {
                DataRow dr = dt.NewRow();
                dr["SEL"] = true;
                dr["Name"] = Name;
                dr["PageCount"] = PageCount.ToUpper().Trim().Replace("P", "");
                dr["URL"] = CombinURL;
                dr["IsDownLoad"] = 0;
                dr["Status"] = "未下載";
                dt.Rows.Add(dr);
            }

            GetliInfo2(NewContents2, ref dt, strURL);

        }
        private string GetTitle(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<div class=\"book-title\">", "</div>");

            GetValue(Value01, out Value02, "<h1>", "</h1>");
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetYEAR(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<strong>出品年代：</strong>", "</span>");
            GetValue(Value01, out Value02, ">", "</a>");
            return Value02;
        }
        private string GetZONE(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<strong>漫畫地區：</strong>", "</span>");
            GetValue(Value01, out Value02, ">", "</a>");
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetLETTER(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<strong>字母索引：</strong>", "</span>");
            GetValue(Value01, out Value02, ">", "</a>");
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetTYPE_SUM(string Contents, out string NewContents)
        {
            string Value01;
            GetValue(Contents, out NewContents, out Value01, "<strong>漫畫劇情：</strong>", "</span>");

            string RESULT = "";
            GetTYPETag(Value01, ref RESULT);

            return RESULT;
        }

        private void GetTYPETag(string Contents, ref string RESULT)
        {
            if (Contents.IndexOf("<a href=\"") < 0) return;
            Contents = Contents.Substring(Contents.IndexOf("<a href=\""));

            string Value02;
            string NewContents;
            GetValue(Contents, out NewContents, out Value02, ">", "</a>");
            RESULT += string.Format("{0}{1}", RESULT.Trim() == "" ? "" : "|", Value02);
            GetTYPETag(NewContents, ref RESULT);

        }

        private string GetWRITER(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<strong>漫畫作者：</strong>", "</span>");
            GetValue(Value01, out Value02, ">", "</a>");
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetNAME2(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<strong>漫畫別名：</strong>", "</span>");
            if (Value01.IndexOf("暫無") > -1)
            {
                Value02 = Value01;
            }
            else
            {
                GetValue(Value01, out Value02, ">", "</a>");
            }
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetSTATUS(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<strong>漫畫狀態：</strong>", "span>");
            GetValue(Value01, out Value02, ">", "</");
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetUPDATE_TIME(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "最近于", "span>");
            GetValue(Value01, out Value02, ">", "</");
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetNEW_PAGE(string Contents, out string NewContents)//格式變化太多 不採用
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "更新至", "a>");
            GetValue(Value01, out Value02, ">", "</");
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetPICTURE(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<p class=\"hcover\">", "alt=");
            GetValue(Value01, out Value02, "\"", "\"");
            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }
        private string GetDOCUMENT(string Contents, out string NewContents)
        {
            string Value01, Value02;
            GetValue(Contents, out NewContents, out Value01, "<div id=\"intro-cut\"", "/div>");

            GetValue(Value01, out Value02, ">", "<");

            Value02 = Value02.Replace("\"", "＂").Replace(",", "，");
            return Value02;
        }

        private List<string> GetAllBook(string strURL)
        {
            List<string> listBOOK = new List<string>();

            DataTable dt = new DataTable();

            dt = GetAllBookDataTable(strURL,"");

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                listBOOK.Add(dt.Rows[i]["Name"].ToString().Trim());
            }

            return listBOOK;
        }
        private DataTable GetAllBookDataTable(string strURL, string Content)
        {
            string Titlt = "";
            return GetAllBookDataTable(strURL, Content, out Titlt);
        }
        private DataTable GetAllBookDataTable(string strURL,string Content,out string titlt)
        {
            titlt = "";
            try
            {
                List<string> listBOOK = new List<string>();

                DataTable dt = new DataTable();
                dt.Columns.Add("SEL");//選取狀態
                dt.Columns.Add("Name");//卷(話)名稱
                dt.Columns.Add("PageCount");//頁數
                dt.Columns.Add("URL");//卷(話)連結
                dt.Columns.Add("SubNum");//代碼
                dt.Columns.Add("IsDownLoad");//是否已下載
                dt.Columns.Add("Status");//狀態
                dt.Columns.Add("DateTime");//狀態

                bool checkAdult = false;
                string strNUM = "";
                string strSUB_NUM = "";
                string P = "";
                GetNumURL(strURL, out strURL, out strNUM, out strSUB_NUM, out P);



                string Contents = "";

                #region 第一種取法 WebClient
                if (Content.Trim() == "")
                {
                    WebClient wc = new WebClient();
                    if (strURL.IndexOf("https:") > -1)//來源是https時 這樣才有辦法讀取
                    {
                        // 強制認為憑證都是通過的，特殊情況再使用
                        ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                        ServicePointManager.SecurityProtocol =
                            SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls |
                            SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                    }
                    byte[] bResult = wc.DownloadData(strURL);
                    Contents = Encoding.UTF8.GetString(bResult);

                }
                else
                    Contents = Content;
                if (Contents.IndexOf("id=\"checkAdult\"") > -1)
                    checkAdult = true;


                string NewContents;
                string Key1 = "<div class=\"book-title\">";
                string Key2 = "</div>";

                //string titlt;
                GetValue(Contents, out NewContents, out titlt, "<h1>", "</h1>");


                getAllLink(strURL, Contents, ref dt);//遞迴取得所有集數的資料 
                #endregion

                #region 第二種取法 WebBrowser
                if (dt == null || dt.Rows.Count < 1)
                {

                    WebBrowser WBSetting = new WebBrowser();
                    try
                    {

                        WBSetting.ScriptErrorsSuppressed = true;
                        //WBSetting.Navigate(string.Format("https://www.ikanman.com/comic/{0}/", strNUM.Trim()));//導頁
                        //WBSetting.Navigate(string.Format("https://www.manhuagui.com/comic/{0}/", strNUM.Trim()));//導頁
                        WBSetting.Navigate(string.Format("{0}", strURL, strNUM.Trim()));//導頁
                        int tryCount = 0;
                        loading(WBSetting, ref tryCount);//等待完成載入
                        if (checkAdult)
                        {
                            checkAdultClick(ref WBSetting);
                        }
                        HtmlDocument doc = WBSetting.Document;

                        var v = from HtmlElement t in doc.All
                                where t.InnerHtml != null && t.TagName.ToUpper().IndexOf("UL") > -1
                                select t;
                        List<HtmlElement> listnewDoc = v.ToList();

                        for (int i = 0; i < listnewDoc.Count; i++)
                        {

                            GetliInfo2(strURL,listnewDoc[i].InnerHtml, ref dt);//依內部資料取得連結
                        }

                    }
                    catch (Exception ex)
                    {

                        return dt.Clone();

                    }

                    WBSetting.Stop();
                    WBSetting.Dispose();
                }
                #endregion


                if (strSUB_NUM.Trim() != "")
                {
                    var v = from t in dt.AsEnumerable()
                            where t["URL"].ToString().Trim().ToUpper().IndexOf("/" + strSUB_NUM.Trim() + ".HTML") > -1
                            select t;
                    if (v != null && v.Count() > 0)
                        dt = v.CopyToDataTable();
                }

                return dt;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void getAllLink(string Contents, ref DataTable dt, string strURL)//遞迴取得所有集數的資料
        {
            if (Contents.IndexOf("<ul") < 0) return;//若未包含則跳過

            string Value01, Value02, NewContents;//所有單行本

            int Index1 = Contents.IndexOf("<ul style=\"display:block\">");
            int Index2 = Contents.IndexOf("<ul>");
            if (Index1 < Index2)
                GetValue(Contents, out NewContents, out Value01, "<ul style=\"display:block\">", "</ul>");//取得內部所有資料
            else
                GetValue(Contents, out NewContents, out Value01, "<ul>", "</ul>");//取得內部所有資料



            GetliInfo(Value01, ref dt, strURL);//依內部資料取得連結
            getAllLink(NewContents, ref dt, strURL);//再繼續挖資料


        }
        #endregion

        private void lvOphen_MouseClick(object sender, MouseEventArgs e)//點最愛時
        {
            if (e.Button == MouseButtons.Left)
            {
                if (((System.Windows.Forms.ListView)sender).Items.Count < 1)//沒點項目 介紹隱藏
                {
                    scpOphen.Panel2Collapsed = true;
                    return;
                }
                //顯示介紹
                ListViewItem Item = ((System.Windows.Forms.ListView)sender).Items[0];
                //NAME , YEAR , ZONE , LETTER , TYPE_SUM , WRITER , NAME2 , STATUS , UPDATE_TIME , NEW_PAGE , DOCUMENT,PICTURE
                labNAME.Text = Item.SubItems[0].Text.ToString().Trim();
                labYEAR.Text = Item.SubItems[1].Text.ToString().Trim();
                labZONE.Text = Item.SubItems[2].Text.ToString().Trim();
                labLETTER.Text = Item.SubItems[3].Text.ToString().Trim();
                labTYPE_SUM.Text = Item.SubItems[4].Text.ToString().Trim();
                labWRITER.Text = Item.SubItems[5].Text.ToString().Trim();
                labNAME2.Text = Item.SubItems[6].Text.ToString().Trim();
                labSTATUS_B.Text = Item.SubItems[7].Text.ToString().Trim();
                labUPDATE_TIME.Text = Item.SubItems[8].Text.ToString().Trim();
                labNEW_PAGE.Text = Item.SubItems[9].Text.ToString().Trim();
                labDOCUMENT.Text = Item.SubItems[10].Text.ToString().Trim();
                if (scpOphen.Panel2Collapsed)//已縮起介紹時 打開
                {
                    scpOphen.Panel2Collapsed = false;
                    //scpOphen.SplitterDistance;
                    //scpOphen.FixedPanel;
                }

                //帶出記錄
            }
            if (e.Button == MouseButtons.Right)
            {
                ListViewItem lvItem = lvOphen.GetItemAt(e.X, e.Y);
                if (lvItem == null) return;
                cmslvOphen.Show(lvOphen, new Point(e.X, e.Y));

            }
        }

        private void lvOphen_MouseDoubleClick(object sender, MouseEventArgs e)//雙點最愛時
        {
            if (e.Button == MouseButtons.Left)
            {
                if (((System.Windows.Forms.ListView)sender).Items.Count < 1) return;
                ListViewItem Item = ((System.Windows.Forms.ListView)sender).Items[0];
                //NAME , YEAR , ZONE , LETTER , TYPE_SUM , WRITER , NAME2 , STATUS , UPDATE_TIME , NEW_PAGE , DOCUMENT,PICTURE,strURL,strNUM
                string strURL = Item.SubItems[12].ToString().Trim();
                //導頁
                if (strURL.Trim() != "")
                {
                    try
                    {
                        tabControl1.SelectedIndex = 0;
                        wbSearch.Stop();
                        wbSearch.Navigate(strURL.Trim());

                    }
                    catch (Exception ex)
                    {

                        wbSearch.Stop();
                    }
                }
                //帶出記錄


            }
        }
        //選最愛項目 事件 沒有選到最愛項目時 介紹隱藏
        private void lvOphen_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(lvOphen.SelectedItems.Count<1)//沒有選到最愛項目時 介紹隱藏
                scpOphen.Panel2Collapsed = true;
        }

        //增加下載記錄
        public void AddRecord(string strURL,string Num,string Name,string SubNum)
        {
            //SubNum是卷(話)的代碼
            if (!CheckURL(strURL)) return;
            //Thread thread = new Thread(AddRecordWork);
            //thread.IsBackground = true;//設為背景執行~才可刪除
            //thread.Start(new object[] { strURL,Num, Name, SubNum, lvRecord, imlRecord, lvRecordBook });//執行 執行緒

            
            //2018-03-06因為多執行緒有時會失敗~所以改單一執行緒
            AddRecordWork(strURL, Num, Name, SubNum, ref lvRecord, ref imlRecord, ref lvRecordBook);


        }
        //檢查URL
        private bool CheckURL(string strURL)
        {
            if (strURL.IndexOf("//tw.manhuagui.com/comic/") < 0 && strURL.IndexOf("//tw.ikanman.com/comic/") < 0
              && strURL.IndexOf("//www.manhuagui.com/") < 0 && strURL.IndexOf("//www.ikanman.com/comic/") < 0)
                return false;
            else
                return true;
        }

        private delegate void delegateAddRecordWork(object obj);
        private void AddRecordWork(object obj)//增加
        {
            if (this.InvokeRequired)//對跨執行緒做處理
            {
                delegateAddRecordWork myUpdate = new delegateAddRecordWork(AddRecordWork);
                this.Invoke(myUpdate, obj);
            }
            else
            {
                int objIndex = 0;
                string strURL = ((object[])obj)[objIndex++].ToString();//傳入的參數
                string Num = ((object[])obj)[objIndex++].ToString();//傳入的參數
                string Name = ((object[])obj)[objIndex++].ToString();//傳入的參數
                string SubNum = ((object[])obj)[objIndex++].ToString();//傳入的參數
                ListView lvRecord_ = (ListView)((object[])obj)[objIndex++];//傳入的參數
                ImageList imlRecord_ = (ImageList)((object[])obj)[objIndex++];//傳入的參數
                ListView lvRecordBook_ = (ListView)((object[])obj)[objIndex++];//傳入的參數
                AddRecordWork(strURL, Num, Name, SubNum, ref lvRecord_, ref imlRecord_,ref lvRecordBook_);
            }
        }

        //增加記錄
        private void AddRecordWork(string strURL,string Num, string Name, string SubNum, ref ListView lvRecord_, ref ImageList imlRecord_,ref ListView lvRecordBook_)
        {
            if (lvRecord_ == null)
                lvRecord_ = this.lvRecord;
            if (imlRecord_ == null)
                imlRecord_ = this.imlRecord;
            if (lvRecordBook_ == null)
                lvRecordBook_ = this.lvRecordBook;

            //一部漫畫一個記錄檔 記錄檔內存所有卷(話)數資料 已下載未下載都記錄

            List<ListViewItem> lvItemItems = new List<ListViewItem>();//有連結存在的項目
           //比對主記錄是否已存在 已存在就跳過建主記錄
           CheckRecordIcon(strURL, Num, SubNum, ref lvRecord_, ref  imlRecord_, ref lvItemItems);


            //從ListView存成CSV檔
            ListViewToCSV(lvRecord_, RecordFileName);

        }

        //比對並新增縮圖項目
        public void CheckRecordIcon(string strURL, string strNUM, string strSUB_NUM, ref ListView lvRecord_, ref ImageList imlRecord_, ref List<ListViewItem> lvItemItems)
        {
            bool IsAddlvRecordItem = true;
            //var vlvItem = from ListViewItem t in lvRecord_.Items
            //              where t.SubItems[12].Text.Trim() == strURL.Trim()
            //              || t.SubItems[12].Text.Trim()+"/" == strURL.Trim()
            //              || t.SubItems[13].Text.Trim() == strNUM.Trim()
            //              select t;
            var vlvItem = from ListViewItem t in this.lvRecord.Items
                          where t.SubItems[12].Text.Trim() == strURL.Trim()
                          || t.SubItems[12].Text.Trim() + "/" == strURL.Trim()
                          || t.SubItems[13].Text.Trim() == strNUM.Trim()
                          select t;
            if (vlvItem != null && vlvItem.Count() > 0)//若已經存在就跳過
            {
                lvItemItems.AddRange(vlvItem);
                IsAddlvRecordItem = false;

            }

            string strDirPath = AppDomain.CurrentDomain.BaseDirectory + "Record\\";
            if (!Directory.Exists(strDirPath))
                Directory.CreateDirectory(strDirPath);
            string FileName = strNUM.Trim() + ".csv";
            string strFilePath = strDirPath + FileName;


            //從CSV取得該漫畫的副記錄資料 CSV=>DataTable
            DataTable dtRecord = new DataTable();//已有的記錄
            GetSettingItem getSettingItem = new GetSettingItem();
            bool IsHaveBookRecord = false;//是否已存在記錄
            bool IsGetBook = true;//每次都去網頁取得比對設true 否則設false
            if (File.Exists(strFilePath))//本機有記錄的話~以本機為主
            {
                dtRecord = getSettingItem.GetCSVTable(strDirPath, FileName);
                IsHaveBookRecord = true;
            }
                

            if (dtRecord == null || dtRecord.Columns.Count < 7) //若無CSV則開新CSV並開欄位
            {
                dtRecord = new DataTable();//已有的記錄
                dtRecord.Columns.Add("SEL");//選取狀態
                dtRecord.Columns.Add("Name");
                dtRecord.Columns.Add("PageCount");
                dtRecord.Columns.Add("URL");
                dtRecord.Columns.Add("SubNum");
                dtRecord.Columns.Add("IsDownLoad");
                dtRecord.Columns.Add("Status");
                dtRecord.Columns.Add("DateTime");
                dtRecord.Columns.Add("Num");

            }
            if (!dtRecord.Columns.Contains("Num"))
                dtRecord.Columns.Add("Num");




            ListViewItem Item = new ListViewItem();
            DataTable dtBOOK = new DataTable();//網頁內所有卷(話)
            GetBookInfoList(strURL, strNUM, ref Item, ref imlRecord_, out dtBOOK, IsHaveBookRecord,IsGetBook);

            if (!IsHaveBookRecord)
                dtBOOK = dtRecord.Clone();
            else
            {

                //dtRecord = dtBOOK.Copy();
            }

            #region 將主記錄 加ITEM
                //將主記錄塞ITEM
            if (IsAddlvRecordItem)
            {
                if (this.lvRecord.InvokeRequired)//在非當前執行緒內 使用委派
                {
                    UpdatelvItemProcess myUpdate = new UpdatelvItemProcess(SetlvItem);
                    lvRecord.Invoke(myUpdate, Item);
                }
                else
                {
                    lvRecord.Items.Add(Item);
                }

                lvItemItems.Add(Item);
            }
            #endregion


            //把本機內副記錄缺少的 由線上副記錄補齊
            dtBOOK.AsEnumerable().ToList().ForEach(X =>
            {
                try
                {
                    string Name = X["Name"].ToString().Trim();
                    var v = from t in dtRecord.AsEnumerable()
                            where t["Name"].ToString().Trim()==Name.Trim()
                            select t;
                    if (v == null || v.Count() < 1)
                    {
                        DataRow dr = dtRecord.NewRow();
                        dr["Name"] = Name.Trim();
                        dr["PageCount"] =X["PageCount"];
                        dr["URL"] =X["URL"].ToString().Trim();
                        dr["SubNum"] = X["SubNum"].ToString().Trim();
                        dr["IsDownLoad"] = X["IsDownLoad"].ToString().Trim();
                        dr["Status"] = X["Status"].ToString().Trim();
                        dr["DateTime"] = X["DateTime"].ToString().Trim();
                        dr["Num"] = strNUM;
                        dtRecord.Rows.Add(dr);
                    }

                }
                catch (Exception ex)
                {

                }

            });

            //補斜線
            if (strURL.LastIndexOf("/") != strURL.Length - 1)
                strURL += "/";


            //從中尋找該卷資料 並標記已下載
            dtRecord.AsEnumerable().ToList().ForEach(X=> {
                if (strSUB_NUM.Trim() != "")//若是是在下載漫畫緒中，會更新該卷(話) 改為已下載 若是是更新網頁資料時不需要跑這段
                {
                    if ((X["URL"].ToString().Trim().ToUpper() == (strURL.Trim().ToUpper() + strSUB_NUM.Trim() + ".HTML") ||
                        X["URL"].ToString().Trim().ToUpper() == strURL.Trim().ToUpper().Replace("TW.", "WWW.") + strSUB_NUM.Trim() + ".HTML")
                       && X["IsDownLoad"].ToString().Trim() != "1")
                    {
                        X["IsDownLoad"] = 1;
                        X["Status"] = "已下載";
                        X["DateTime"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    else if (X["IsDownLoad"].ToString().Trim() == "")
                    {
                        X["IsDownLoad"] = 0;
                        X["Status"] = "未下載";
                    }
                }

                X["Num"] = strNUM;
            });

            if (dtRecord.Columns.Contains("SEL"))
                dtRecord.Columns.Remove("SEL");
            //將DataTable存成CSV
            getSettingItem.SetCSVTable(dtRecord, strDirPath, FileName);

            Thread.Sleep(100);





        }
        //點記錄項目時
        private void lvRecord_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (((System.Windows.Forms.ListView)sender).Items.Count < 1)//沒點項目 介紹子記錄
                {
                    scRecord.Panel2Collapsed = true;
                    return;
                }

                if (scRecord.Panel2Collapsed)//已縮起介紹時 打開
                {
                    scRecord.Panel2Collapsed = false;
                    //scpOphen.SplitterDistance;
                    //scpOphen.FixedPanel;
                }

                //帶出記錄
                lvRecordBook.Items.Clear();//先清空
                
                //顯示詳細記錄
                ListViewItem lvRecordItem = ((System.Windows.Forms.ListView)sender).SelectedItems[0];
                string strNUM = lvRecordItem.SubItems[13].Text.Trim();
                string MainName = lvRecordItem.SubItems[0].Text.ToString().Trim();
                //取得該漫畫的記錄資料 CSV=>DataTable
                string strDirPath = AppDomain.CurrentDomain.BaseDirectory + "Record\\";
                if (!Directory.Exists(strDirPath))
                    Directory.CreateDirectory(strDirPath);
                string FileName = strNUM.Trim() + ".csv";
                string strFilePath = strDirPath + FileName;
                DataTable dtRecord = new DataTable();//已有的記錄
                GetSettingItem getSettingItem = new GetSettingItem();
                if (File.Exists(strFilePath))
                {
                    dtRecord = getSettingItem.GetCSVTable(strDirPath, FileName);
                }



                for (int i = 0; i < dtRecord.Rows.Count; i++)
                {
                    DataRow dr = dtRecord.Rows[i];

                    //ListView增加項目
                    ListViewItem Item = new ListViewItem(new string[] {
                    dr["Name"].ToString().Trim() ,
                    dr["PageCount"].ToString().Trim() ,
                    dr["URL"].ToString().Trim() ,
                    dr["SubNum"].ToString().Trim() ,
                    dr["IsDownLoad"].ToString().Trim() ,
                    dr["Status"].ToString().Trim() ,
                    dr["DateTime"].ToString().Trim(),
                    MainName
                });
                    Item.Text = dr["NAME"].ToString().Trim();
                    Item.ToolTipText = dr["URL"].ToString().Trim();
                    string strURL = dr["URL"].ToString().Trim();
                    strURL += strURL.LastIndexOf("/") == strURL.Length - 1 ? "" : "/";
                    Item.Tag = strURL.Trim();
                    if (lvRecordBook.InvokeRequired)//在非當前執行緒內 使用委派
                    {
                        UpdatelvItemProcess myUpdate = new UpdatelvItemProcess(SetlvItem);
                        lvRecordBook.Invoke(myUpdate, lvRecordBook, Item);
                    }
                    else
                    {
                        lvRecordBook.Items.Add(Item);
                    }



                    //Thread.Sleep(100);
                }

            }
            if (e.Button == MouseButtons.Right)
            {
                ListViewItem lvItem = lvRecord.GetItemAt(e.X, e.Y);
                if (lvItem == null) return;
                cmsRecord.Show(lvRecord, new Point(e.X, e.Y));

            }
        }

        private void lvRecord_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (((System.Windows.Forms.ListView)sender).Items.Count < 1) return;
                ListViewItem Item = ((System.Windows.Forms.ListView)sender).Items[0];
                //NAME , YEAR , ZONE , LETTER , TYPE_SUM , WRITER , NAME2 , STATUS , UPDATE_TIME , NEW_PAGE , DOCUMENT,PICTURE,strURL,strNUM
                string strURL = Item.SubItems[12].ToString().Trim();
                //導頁
                if (strURL.Trim() != "")
                {
                    try
                    {
                        tabControl1.SelectedIndex = 0;
                        wbSearch.Stop();
                        wbSearch.Navigate(strURL);

                    }
                    catch (Exception ex)
                    {

                        wbSearch.Stop();
                    }
                }

            }
        }

        //選記錄項目 事件 沒有選到記錄項目時 介紹副記錄
        private void lvRecord_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvRecord.SelectedItems.Count < 1)//沒有選到最愛項目時 介紹隱藏
                scRecord.Panel2Collapsed = true;
        }

        private void lvRecordBook_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ListViewItem lvItem = lvRecordBook.GetItemAt(e.X, e.Y);
                if (lvItem == null) return;
                cmslvRecordBook.Show(lvRecordBook, new Point(e.X, e.Y));

            }
        }
        private void lvRecordBook_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            if (lvRecordBook.SelectedItems.Count < 1) return;
            //下載
            string strURL = lvRecordBook.SelectedItems[0].SubItems[2].Text.Trim();
            if (strURL.Trim() == "") return;
            txtURL.Text = strURL;
            btnSetting.PerformClick();
            btnGetBySetting.PerformClick();
        }
        //勾選進行 下載記錄
        private void chkRecord_CheckedChanged(object sender, EventArgs e)
        {
            GetSettingItem g = new GetSettingItem();
            g.SetValue("IsRecord", chkRecord.Checked?"True":"False");
        }

        #region 右鍵目錄
        //選擇下載
        private void cmsOphen_Download_Click(object sender, EventArgs e)
        {
            if (lvOphen.SelectedItems.Count < 1) return;

            string strURL = lvOphen.SelectedItems[0].SubItems[12].Text.Trim();
            if (strURL.Trim() == "") return;
            txtURL.Text = strURL;
            btnSetting.PerformClick();
        }
        //下載全部
        private void cmsOphen_DownloadAll_Click(object sender, EventArgs e)
        {
            if (lvOphen.SelectedItems.Count < 1) return;

            string strURL = lvOphen.SelectedItems[0].SubItems[12].Text.Trim();
            if (strURL.Trim() == "") return;
            txtURL.Text = strURL;
            btnSetting.PerformClick();
            btnGetBySetting.PerformClick();
        }
        //導頁
        private void cmsOphen_Redirect_Click(object sender, EventArgs e)
        {
            if (lvOphen.SelectedItems.Count < 1) return;

            string strURL = lvOphen.SelectedItems[0].SubItems[12].Text.Trim();
            if (strURL.Trim() == "") return;

            try
            {
                tabControl1.SelectedIndex = 0;
                wbSearch.Stop();
                wbSearch.Navigate(strURL);

            }
            catch (Exception ex)
            {

                wbSearch.Stop();
            }

        }
        //移除
        private void cmsOphen_Delete_Click(object sender, EventArgs e)
        {
            if (lvOphen.SelectedItems.Count < 1) return;
            DialogResult Resualt = MessageBox.Show("確定移除?", "移除", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            if (Resualt == DialogResult.Cancel)
                return;
            for (int i = 0; i < lvOphen.SelectedItems.Count; i++)
            {
                lvOphen.Items.Remove(lvOphen.SelectedItems[i]);
            }
            ListViewToCSV(lvOphen, OphenFileName);
            //DELETE PICTURE

        }
        //查看記錄
        private void cmsOphen_Record_Click(object sender, EventArgs e)
        {
            if (lvOphen.SelectedItems.Count < 1) return;

            string strURL_ = lvOphen.SelectedItems[0].SubItems[12].Text.Trim();
            if (strURL_.Trim() == "") return;

            //帶出記錄
            lvRecordBook.Items.Clear();//先清空
            var v = from ListViewItem t in lvRecord.Items
                    where t.SubItems[12].Text.Trim() == strURL_
                    select t;
            if (v == null || v.Count() < 1) return;

            //顯示詳細記錄
            ListViewItem lvRecordItem = (ListViewItem)v.ToList()[0];
            lvRecordItem.Selected = true;
            string strNUM = lvRecordItem.SubItems[13].Text.ToString().Trim();

            //取得該漫畫的記錄資料 CSV=>DataTable
            string strDirPath = AppDomain.CurrentDomain.BaseDirectory + "Record\\";
            if (!Directory.Exists(strDirPath))
                Directory.CreateDirectory(strDirPath);
            string FileName = strNUM.Trim() + ".csv";
            string strFilePath = strDirPath + FileName;
            DataTable dtRecord = new DataTable();//已有的記錄
            GetSettingItem getSettingItem = new GetSettingItem();
            if (File.Exists(strFilePath))
            {
                dtRecord = getSettingItem.GetCSVTable(strDirPath, FileName);
            }

            tabControl1.SelectedIndex = 1;

            if (scRecord.Panel2Collapsed)//已縮起介紹時 打開
            {
                scRecord.Panel2Collapsed = false;
                //scpOphen.SplitterDistance;
                //scpOphen.FixedPanel;
            }
            for (int i = 0; i < dtRecord.Rows.Count; i++)
            {
                DataRow dr = dtRecord.Rows[i];

                //ListView增加項目
                ListViewItem Item = new ListViewItem(new string[] {
                    dr["Name"].ToString().Trim() ,
                    dr["PageCount"].ToString().Trim() ,
                    dr["URL"].ToString().Trim() ,
                    dr["SubNum"].ToString().Trim() ,
                    dr["IsDownLoad"].ToString().Trim() ,
                    dr["Status"].ToString().Trim() ,
                    dr["DateTime"].ToString().Trim()
                });
                Item.Text = dr["NAME"].ToString().Trim();
                Item.ToolTipText = dr["URL"].ToString().Trim();
                string strURL = dr["URL"].ToString().Trim();
                strURL += strURL.LastIndexOf("/") == strURL.Length - 1 ? "" : "/";
                Item.Tag = strURL.Trim();
                if (lvRecordBook.InvokeRequired)//在非當前執行緒內 使用委派
                {
                    UpdatelvItemProcess myUpdate = new UpdatelvItemProcess(SetlvItem);
                    lvRecordBook.Invoke(myUpdate, lvRecordBook, Item);
                }
                else
                {
                    lvRecordBook.Items.Add(Item);
                }



                Thread.Sleep(100);
            }


        }

        //開啟資料夾
        private void cmsOphen_Dir_Click(object sender, EventArgs e)
        {
            if (lvOphen.SelectedItems.Count < 1) return;
            if (!chkIsUseDefaultSavePath.Checked) return;
            string DirRootPath = txtDirPath.Text.Trim();
            if (DirRootPath.LastIndexOf("\\") != DirRootPath.Length - 1)
                DirRootPath += "\\";

            string Name = lvOphen.SelectedItems[0].SubItems[0].Text.Trim();

            string DirPath = DirRootPath + Name;

            if (!Directory.Exists(DirPath))
                Directory.CreateDirectory(DirPath);

            System.Diagnostics.Process.Start(DirPath);
        }

        //下載
        private void cmsRecordBook_Download_Click(object sender, EventArgs e)
        {
            if (lvRecordBook.SelectedItems.Count < 1) return;
            for (int i = 0; i < lvRecordBook.SelectedItems.Count; i++)
            {
                string strURL = lvRecordBook.SelectedItems[i].SubItems[2].Text.Trim();
                if (strURL.Trim() == "") return;
                txtURL.Text = strURL;
                btnSetting.PerformClick();
                btnGetBySetting.PerformClick();
            }
        }
        //導頁
        private void cmsRecordBook_Redirect_Click(object sender, EventArgs e)
        {
            if (lvRecordBook.SelectedItems.Count < 1) return;

            string strURL = lvRecordBook.SelectedItems[0].SubItems[2].Text.Trim();
            if (strURL.Trim() == "") return;
            try
            {
                tabControl1.SelectedIndex = 0;
                wbSearch.Stop();
                wbSearch.Navigate(strURL);

            }
            catch (Exception ex)
            {

                wbSearch.Stop();
            }
        }

        //標示「未下載」
        private void cmsRecordBook_UnDownload_Click(object sender, EventArgs e)
        {
            if (lvRecordBook.SelectedItems.Count < 1) return;
            for (int i = 0; i < lvRecordBook.SelectedItems.Count; i++)
            {
                lvRecordBook.SelectedItems[i].SubItems[4].Text = "0";
                lvRecordBook.SelectedItems[i].SubItems[5].Text = "未下載";
            }
            string strURL = lvRecordBook.SelectedItems[0].SubItems[2].Text.Trim();
            string strNUM = "";
            string strSUB_NUM = "";
            string P = "";
            GetNumURL(strURL, out strURL, out strNUM, out strSUB_NUM, out P);

            if (strNUM.Trim() == "") return;
            string strDirPath = AppDomain.CurrentDomain.BaseDirectory + "Record\\";
            if (!Directory.Exists(strDirPath))
                Directory.CreateDirectory(strDirPath);
            string FileName = strNUM.Trim() + ".csv";
            string strFilePath = strDirPath + FileName;

            lvRecordBookToCSV(lvRecordBook, strDirPath, FileName);
        }
        //標示「已下載」
        private void cmsRecordBook_HaveDownload_Click(object sender, EventArgs e)
        {
            if (lvRecordBook.SelectedItems.Count < 1) return;
            for (int i = 0; i < lvRecordBook.SelectedItems.Count; i++)
            {
                lvRecordBook.SelectedItems[i].SubItems[4].Text = "1";
                lvRecordBook.SelectedItems[i].SubItems[5].Text = "已下載";
            }
            string strURL = lvRecordBook.SelectedItems[0].SubItems[2].Text.Trim();
            string strNUM = "";
            string strSUB_NUM = "";
            string P = "";
            GetNumURL(strURL, out strURL, out strNUM, out strSUB_NUM, out P);
            if (strNUM.Trim() == "") return;
            string strDirPath = AppDomain.CurrentDomain.BaseDirectory + "Record\\";
            if (!Directory.Exists(strDirPath))
                Directory.CreateDirectory(strDirPath);
            string FileName = strNUM.Trim() + ".csv";
            string strFilePath = strDirPath + FileName;
            lvRecordBookToCSV(lvRecordBook, strDirPath, FileName);
        }
        //開啟資料夾
        private void cmsRecordBook_Dir(object sender, EventArgs e)
        {
            if (lvRecordBook.SelectedItems.Count < 1) return;
            if (!chkIsUseDefaultSavePath.Checked) return;
            string DirRootPath = txtDirPath.Text.Trim();
            if (DirRootPath.LastIndexOf("\\") != DirRootPath.Length - 1)
                DirRootPath += "\\";

            string Name = lvRecordBook.SelectedItems[0].SubItems[0].Text.Trim();
            string PageCount = lvRecordBook.SelectedItems[0].SubItems[1].Text.Trim();
            string MainName = lvRecordBook.SelectedItems[0].SubItems[7].Text.Trim();

            string DirPath = string.Format("{0}{1}\\{2}[{3}p]",DirRootPath , MainName ,Name,PageCount);

            if (!Directory.Exists(DirPath))
                Directory.CreateDirectory(DirPath);

            System.Diagnostics.Process.Start(DirPath);
        }


        //選擇下載
        private void cmsRecord_Download_Click(object sender, EventArgs e)
        {
            if (lvRecord.SelectedItems.Count < 1) return;

            string strURL = lvRecord.SelectedItems[0].SubItems[12].Text.Trim();
            if (strURL.Trim() == "") return;
            txtURL.Text = strURL;
            btnSetting.PerformClick();

        }
        //下載全部
        private void cmsRecord_DownloadAll_Click(object sender, EventArgs e)
        {
            if (lvRecord.SelectedItems.Count < 1) return;

            string strURL = lvRecord.SelectedItems[0].SubItems[12].Text.Trim();
            if (strURL.Trim() == "") return;
            txtURL.Text = strURL;
            btnSetting.PerformClick();
            btnGetBySetting.PerformClick();
        }
        //導頁
        private void cmsRecord_Redirect_Click(object sender, EventArgs e)
        {
            if (lvRecord.SelectedItems.Count < 1) return;

            string strURL = lvRecord.SelectedItems[0].SubItems[12].Text.Trim();
            if (strURL.Trim() == "") return;
            try
            {
                tabControl1.SelectedIndex = 0;
                wbSearch.Stop();
                wbSearch.Navigate(strURL);

            }
            catch (Exception ex)
            {

                wbSearch.Stop();
            }
        }
        //移除
        private void cmsRecord_Delete_Click(object sender, EventArgs e)
        {
            if (lvRecord.SelectedItems.Count < 1) return;
            DialogResult Resualt = MessageBox.Show("確定移除?", "移除", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            if (Resualt == DialogResult.Cancel)
                return;

            for (int i = 0; i < lvRecord.SelectedItems.Count; i++)
            {
                lvRecord.Items.Remove(lvRecord.SelectedItems[i]);
            }
            ListViewToCSV(lvRecord, RecordFileName);

        }

        //開啟資料夾
        private void cmsRecord_Dir_Click(object sender, EventArgs e)
        {
            if (lvRecord.SelectedItems.Count < 1) return;
            if (!chkIsUseDefaultSavePath.Checked) return;
            string DirRootPath = txtDirPath.Text.Trim();
            if (DirRootPath.LastIndexOf("\\") != DirRootPath.Length - 1)
                DirRootPath += "\\";

            string Name = lvRecord.SelectedItems[0].SubItems[0].Text.Trim();

            string DirPath = DirRootPath + Name;

            if (!Directory.Exists(DirPath))
                Directory.CreateDirectory(DirPath);

            System.Diagnostics.Process.Start(DirPath);
        }
        #endregion
        //工作執行數(目前不啟用 因為還未使用多執行緒)
        private void cmbWorkCount_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbWorkCount.SelectedItem.ToString().Trim() == "") return;
            

            ShowWorkList();
        }


        #region ========加密========
        /// <summary>
        /// SHA256加密
        /// </summary>
        /// <param name="source">來源字串</param>
        /// <returns>加密字串</returns>
        public string AesEncrypt(string source)
        {
            return AesEncrypt(source, "我跟你說你不要跟別人說，你若跟別人說，不要跟別人說是我叫你不要跟別人說");
        }


        /// <summary>
        /// SHA256加密
        /// </summary>
        /// <param name="source">來源字串</param>
        /// <param name="strKey">加密Key</param>
        /// <returns>加密字串</returns>
        public string AesEncrypt(string source, string strKey)
        {
            StringBuilder sb = new StringBuilder();
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            SHA256CryptoServiceProvider sha256 = new SHA256CryptoServiceProvider();
            AesCryptoServiceProvider provider = new AesCryptoServiceProvider();

            byte[] keyData = sha256.ComputeHash(Encoding.UTF8.GetBytes(strKey));
            byte[] IVData = md5.ComputeHash(Encoding.UTF8.GetBytes(strKey));

            provider.Key = keyData;
            provider.IV = IVData;

            //byte[] key = Encoding.ASCII.GetBytes("12345678");
            //byte[] iv = Encoding.ASCII.GetBytes("87654321");
            byte[] dataByteArray = Encoding.UTF8.GetBytes(source);


            string encrypt = "";
            using (MemoryStream ms = new MemoryStream())
            using (CryptoStream cs = new CryptoStream(ms, provider.CreateEncryptor(), CryptoStreamMode.Write))
            {
                cs.Write(dataByteArray, 0, dataByteArray.Length);
                cs.FlushFinalBlock();
                //輸出資料
                foreach (byte b in ms.ToArray())
                {
                    sb.AppendFormat("{0:X2}", b);
                }
                encrypt = sb.ToString();
            }
            return encrypt;
        }
        #endregion
        #region ========解密========
        /// <summary>
        /// SHA256解密
        /// </summary>
        /// <param name="encrypt">來源字串</param>
        /// <returns>解密字串</returns>
        public string AesDecrypt(string encrypt)
        {
            return AesDecrypt(encrypt, "我跟你說你不要跟別人說，你若跟別人說，不要跟別人說是我叫你不要跟別人說");
        }

        int currentCol = -1;
        bool sort = true;
        private void lvRecordBook_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            string Asc = ((char)0x25bc).ToString().PadLeft(4, ' ');
            string Des = ((char)0x25b2).ToString().PadLeft(4, ' ');

            if (sort == false)
            {
                sort = true;
                string oldStr = this.lvRecordBook.Columns[e.Column].Text.TrimEnd((char)0x25bc, (char)0x25b2, ' ');
                this.lvRecordBook.Columns[e.Column].Text = oldStr + Des;
            }
            else if (sort == true)
            {
                sort = false;
                string oldStr = this.lvRecordBook.Columns[e.Column].Text.TrimEnd((char)0x25bc, (char)0x25b2, ' ');
                this.lvRecordBook.Columns[e.Column].Text = oldStr + Asc;
            }

            if (e.Column == 1)//在設計器中把列頭的tag設為"n",則表示該列按數字比較器處理,否則為文本
                lvRecordBook.ListViewItemSorter = new ListViewItemComparerNum(e.Column, sort);
            else
                lvRecordBook.ListViewItemSorter = new ListViewItemComparer(e.Column, sort);
            this.lvRecordBook.Sort();
            int rowCount = this.lvRecordBook.Items.Count;
            if (currentCol != -1)
            {
                if (e.Column != currentCol)
                    this.lvRecordBook.Columns[currentCol].Text = this.lvRecordBook.Columns[currentCol].Text.TrimEnd((char)0x25bc, (char)0x25b2, ' ');
            }
            currentCol = e.Column;
        }

      
        //更新網頁資料
        private void cmsRecordBook_UpdateRecordItem_Click(object sender, EventArgs e)
        {
            if (lvRecordBook.SelectedItems.Count < 1) return;
            string strSubNum = lvRecordBook.SelectedItems[0].SubItems[3].Text.Trim();
            if (strSubNum.Trim() == "") return;
            try
            {
                List<ListViewItem> lvItemItems = new List<ListViewItem>();//有連結存在的項目
                //CheckRecordIcon(strURL, "Num", "", ref lvRecord, ref imlRecord, ref lvItemItems);

            }
            catch (Exception ex)
            {

      
            }


            
        }
        //更新網頁資料
        private void cmsRecord_Update_Click(object sender, EventArgs e)
        {
            if (lvRecord.SelectedItems.Count < 1) return;

            string strNum = lvRecord.SelectedItems[0].SubItems[13].Text.Trim();
            string strURL = lvRecord.SelectedItems[0].SubItems[12].Text.Trim();

            List<ListViewItem> lvItemItems = new List<ListViewItem>();//有連結存在的項目
            CheckRecordIcon(strURL, strNum, "", ref lvRecord, ref imlRecord, ref lvItemItems);
        }

        List<int> listSelTemp = new List<int>();
        List<int> listSelTemp2 = new List<int>();
        private void gvList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            listSelTemp = new List<int>(listSelTemp2);
            listSelTemp2.Clear();
            if (e.ColumnIndex == -1|| e.ColumnIndex == 0)
            {
                for (int i = 0; i < gvList.Rows.Count; i++)
                {
                    if (gvList.Rows[i].Cells[0].Selected)
                    {
                        gvList.Rows[i].Selected = true;
                        listSelTemp2.Add(i);
                    }
                    else
                        gvList.Rows[i].Selected = false;
                }
            }

                //DataTable dt = (DataTable)gvList.DataSource;
                //dt.AsEnumerable().ToList().ForEach(X => X["SEL"] = "True");
                
            //for (int i = 0; i < gvList.Rows.Count; i++)
            //{
            //    if (gvList.Rows[i].Selected)
            //        listSelTemp.Add(i);
            //}

        }



        private void gvList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //非選擇第一欄時，跳過
            if (e.ColumnIndex != 0) return;

            string SEL = gvList.Rows[e.RowIndex].Cells[0].Value.ToString().Trim()=="True"?"False":"True";

            listSelTemp2.Clear();

            for (int i = 0; i < gvList.Rows.Count; i++)
            {
                //if (e.RowIndex == i) continue;

                var v = from t in listSelTemp
                        where t == i
                        select t;
                if (v != null && v.Count() > 0)
                {
                    gvList.Rows[i].Cells[0].Value = SEL;
                    gvList.Rows[i].Selected = true;
                    listSelTemp2.Add(i);
                }

            }
        }

        



        /// <summary>
        /// SHA256解密
        /// </summary>
        /// <param name="encrypt">來源字串</param>
        /// <param name="strKey">解密Key</param>
        /// <returns>解密字串</returns>
        public string AesDecrypt(string encrypt, string strKey)
        {
            byte[] dataByteArray = new byte[encrypt.Length / 2];
            for (int x = 0; x < encrypt.Length / 2; x++)
            {
                int i = (Convert.ToInt32(encrypt.Substring(x * 2, 2), 16));
                dataByteArray[x] = (byte)i;
            }

            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            SHA256CryptoServiceProvider sha256 = new SHA256CryptoServiceProvider();
            AesCryptoServiceProvider provider = new AesCryptoServiceProvider();

            byte[] keyData = sha256.ComputeHash(Encoding.UTF8.GetBytes(strKey));
            byte[] IVData = md5.ComputeHash(Encoding.UTF8.GetBytes(strKey));

            provider.Key = keyData;
            provider.IV = IVData;

            using (MemoryStream ms = new MemoryStream())
            {
                using (CryptoStream cs = new CryptoStream(ms, provider.CreateDecryptor(), CryptoStreamMode.Write))
                {
                    cs.Write(dataByteArray, 0, dataByteArray.Length);
                    cs.FlushFinalBlock();
                    return Encoding.UTF8.GetString(ms.ToArray());
                }
            }
        }
        #endregion
    }

    public class WorkOption
    {
        public WorkOption(Thread _thread,int _NO, string _ID, int _IsWork, int _Status, WebBrowser _WB, string _URL, int _PageIndex, int _PageCount, string _DirPath, string _MainName, string _Name, string _Message, int _Width, int _Height, int _SaveProcessType, int _RetryCnt,ToolStripProgressBar _pbOne,int _NavigateFlag)
        {
            this.thread = _thread;
            //this.thread.SetApartmentState(ApartmentState.STA);
            this.ID = _ID;
            this.IsWork = _IsWork;
            this.Status = _Status;
            this.WB = _WB;
            this.URL = _URL;
            this.PageIndex = _PageIndex;
            this.PageCount = _PageCount;
            this.DirPath = _DirPath;
            this.MainName = _MainName;
            this.Name = _Name;
            this.Message = _Message;
            this.Width = _Width;
            this.Height = _Height;
            this.SaveProcessType = _SaveProcessType;
            this.RetryCnt = _RetryCnt;
            this.pbOne = _pbOne;
            this.NavigateCNT = 0;
            this.NavigateFlag = _NavigateFlag;
            this.NO = _NO;
            
        }

        public Thread thread { get; set; }//執行緒
        public int NO { get; set; }//序號

        string _ID;
        public string ID//給一組唯一的編號
        {
            get { return _ID; }

            set
            {
                _ID = value;
                this.btnOpenDir.Tag = value;
                this.btnStop.Tag = value;
                this.btnDel.Tag = value;
            }
        }
 
        public int IsWork { get; set; }//0未 1已啟動 2暫停中
        int _Status;
        public int Status//0初始 1執行中 2已完成 3暫停中 4錯誤停止
        {
            get { return _Status; }

            set
            {
                _Status = value;

                if (value == 3)
                    this.btnStop.Text = "繼續";
                else
                    this.btnStop.Text = "暫停";

                if (value != 4)
                    this.labStatus.Text = value == 0 ? "未執行" : value == 1 ? "進行中" : value == 2 ? "已完成" : value == 3 ? "暫停中" : value == 4 ? "錯誤停止" : "不明";
            }
        }

        int _PageIndex=0;
        public int PageIndex//目前在第幾頁 啟始值為1
        {
            get { return _PageIndex; }

            set
            {
                _PageIndex = value;
                string WorkCount = string.Format("{0}/{1}", value.ToString(), this.PageCount.ToString()); //執行數量狀況
                this.labWorkCount.Text = WorkCount;

                if(this.pbOne!=null && this.PageCount !=0)
                {
                    this.pbOne.Maximum = this.PageCount;
                    this.pbOne.Value = value<=this.PageCount? value: this.PageCount;
                }
            }
        }

        int _PageCount=0;
        public int PageCount//總頁數
        {
            get { return _PageCount; }

            set
            {
                _PageCount = value;
                string WorkCount = string.Format("{0}/{1}", this.PageIndex.ToString(), value.ToString()); //執行數量狀況
                this.labWorkCount.Text = WorkCount;
            }
        }

        string _Message;
        public string Message//訊息
        {
            get { return _Message; }

            set
            {
                _Message = value;
                this.labStatus2.Text = value;//訊息

            }
        }

        public int Width { get; set; }//寬度
        public int Height { get; set; }//高度

        public WebBrowser WB { get; set; }//瀏覽器
        public string URL { get; set; }//網址

        string _DirPath;
        public string DirPath//儲存的主資料夾位置
        {
            get { return _DirPath.Replace("/", "").Replace("*", "").Replace("?", "").Replace(">", "").Replace("<", "").Replace("|", ""); }
            set{_DirPath = value;}
        }

        //.Replace("\\","").Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace(">", "").Replace("<", "").Replace("|", "")
        string _MainName;
        public string MainName//漫畫名稱
        {
            get { return _MainName.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace(">", "").Replace("<", "").Replace("|", ""); }

            set
            {
                _MainName = value;
                string Temp = value;
                if (value.Length > 5)//長度過長 截掉
                    Temp = value.Substring(0, 3) + "...";
                btnOpenDir.Text = string.Format("[{0}-{1}]", Temp, this.Name??"");               
            }
        }

        string _Name;
        public string Name//單行本或是該連載集數名稱
        {
            get { return _Name!=null?_Name.Replace("\\", "").Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace(">", "").Replace("<", "").Replace("|", ""):null; }

            set
            {
                _Name = value;
                string Temp = MainName??"";
                if (value.Length > 5)//長度過長 截掉
                    Temp = value.Substring(0, 3) + "...";
                btnOpenDir.Text = string.Format("[{0}-{1}]", Temp, value);
            }
        }

        public int SaveProcessType { get; set; }//內頁重覆的處理方式 0跳過 1覆蓋 2重新命名

        public int RetryCnt { get; set; }//重試次數

        public int NavigateCNT { get; set; }//導頁次數
        public int NavigateFlag { get; set; }//導頁次數限制

        CheckBox _chkSEL;
        public CheckBox chkSEL//選取
        {
            get
            {
                if (_chkSEL == null)
                    _chkSEL = new CheckBox();
                return _chkSEL;
            }

            set
            {
                if (_chkSEL == null)
                    _chkSEL = new CheckBox();
                _chkSEL = value;
            }
        }

        Label _labNo;
        public Label labNo//序號
        {
            get
            {
                if (_labNo == null)
                    _labNo = new Label();
                return _labNo;
            }

            set
            {
                if (_labNo == null)
                    _labNo = new Label();
                _labNo = value;
            }
        }

        Label _labStatus;
        public Label labStatus//狀態
        {
            get
            {
                if (_labStatus == null)
                    _labStatus = new Label();
                return _labStatus;
            }

            set
            {
                if (_labStatus == null)
                    _labStatus = new Label();
                _labStatus = value;
            }
        }

        Button _btnOpenDir;
        public Button btnOpenDir//卷名及開啟資料夾
        {
            get
            {
                if (_btnOpenDir == null)
                    _btnOpenDir = new Button();
                return _btnOpenDir;
            }

            set
            {
                if (_btnOpenDir == null)
                    _btnOpenDir = new Button();
                _btnOpenDir = value;
            }
        }

        Label _labWorkCount;
        public Label labWorkCount//執行數量狀況
        {
            get
            {
                if (_labWorkCount == null)
                    _labWorkCount = new Label();
                return _labWorkCount;
            }

            set
            {
                if (_labWorkCount == null)
                    _labWorkCount = new Label();
                _labWorkCount = value;
            }
        }

        Label _labStatus2;
        public Label labStatus2//訊息
        {
            get
            {
                if (_labStatus2 == null)
                    _labStatus2 = new Label();
                return _labStatus2;
            }

            set
            {
                if (_labStatus2 == null)
                    _labStatus2 = new Label();
                _labStatus2 = value;
            }
        }

        Button _btnMain;
        public Button btnMain//漫畫連結
        {
            get
            {
                if (_btnMain == null)
                    _btnMain = new Button();
                _btnMain.Text = "主連結";
                return _btnMain;
            }

            set
            {
                if (_btnMain == null)
                    _btnMain = new Button();
                _btnMain = value;
            }
        }

        Button _btnName;
        public Button btnName//卷連結
        {
            get
            {
                if (_btnName == null)
                    _btnName = new Button();
                _btnName.Text = "內連結";
                return _btnName;
            }

            set
            {
                if (_btnName == null)
                    _btnName = new Button();
                _btnName = value;
            }
        }
        Button _btnSkip;
        public Button btnSkip//略過一頁
        {
            get
            {
                if (_btnSkip == null)
                    _btnSkip = new Button();
                _btnSkip.Text = "跳過";
                return _btnSkip;
            }

            set
            {
                if (_btnSkip == null)
                    _btnSkip = new Button();
                _btnSkip = value;
            }
        }

        Button _btnStop;
        public Button btnStop//暫停
        {
            get
            {
                if (_btnStop == null)
                    _btnStop = new Button();
                return _btnStop;
            }

            set
            {
                if (_btnStop == null)
                    _btnStop = new Button();
                _btnStop = value;
            }
        }

        Button _btnDel;
        public Button btnDel//刪除
        {
            get
            {
                if (_btnDel == null)
                    _btnDel = new Button();
                return _btnDel;
            }

            set
            {
                if (_btnDel == null)
                    _btnDel = new Button();
                _btnDel = value;
            }
        }

        public ToolStripProgressBar pbOne{ get; set; }//單一工作進度表

        public FlowLayoutPanel FLP { get; set; }//單一工作列

        public bool IsSleep { get; set; }//是否需要下載延遲功能

        public int SleepMilliseconds { get; set; }//延遲秒數



    }
    public class GetSettingItem
    {
        public GetSettingItem()
        {

        }

        static string striniFilePath;

        public GetSettingItem(string strPath)
        {
            striniFilePath = strPath;
        }

        public string GetSetValue(string strKey)
        {
            try
            {
                string strValue = "";
                string strPath = "./Setting.ini";
                if (striniFilePath != null && striniFilePath.Trim() != "")
                    strPath = striniFilePath + "\\Setting.ini";
                else
                    strPath = System.AppDomain.CurrentDomain.BaseDirectory + "\\Setting.ini";
                FileInfo F = new FileInfo(strPath);
                string[] strValues = File.ReadLines(F.FullName).ToArray();
                foreach (string strLine in strValues)
                {
                    if (strLine.IndexOf("<add key=\"") < 0) continue;
                    int start = strLine.IndexOf("<add key=\"") + 10;
                    string Temp = strLine.Substring(start);
                    int end = Temp.IndexOf("\"");
                    if (strKey.Trim().ToUpper() == Temp.Substring(0, end).Trim().ToUpper())
                    {
                        int start2 = Temp.IndexOf("value=\"") + 7;
                        strValue = Temp.Substring(start2, Temp.LastIndexOf("\"") - start2);
                    }
                }

                return strValue;
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        public string GetSetValue(string strPath, string strKey)
        {
            try
            {
                string strValue = "";
                //string strPath = "./Setting.ini";
                strPath = string.Format("{0}/{1}", strPath, "Setting.ini");
                FileInfo F = new FileInfo(strPath);
                string[] strValues = File.ReadLines(F.FullName).ToArray();
                foreach (string strLine in strValues)
                {
                    if (strLine.IndexOf("<add key=\"") < 0) continue;
                    int start = strLine.IndexOf("<add key=\"") + 10;
                    string Temp = strLine.Substring(start);
                    int end = Temp.IndexOf("\"");
                    if (strKey.Trim().ToUpper() == Temp.Substring(0, end).Trim().ToUpper())
                    {
                        int start2 = Temp.IndexOf("value=\"") + 7;
                        strValue = Temp.Substring(start2, Temp.LastIndexOf("\"") - start2);
                    }
                }

                return strValue;
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        public bool SetValue(string strKey, string Value)
        {
            return SetValue("", strKey, Value);
        }

        public bool SetValue(string strPath, string strKey, string Value)
        {

            try
            {
                if (strPath.Trim() == "")
                    strPath = GetFilePath("", "./Setting.ini");
                FileInfo F = new FileInfo(strPath);
                string[] strValues = File.ReadLines(F.FullName).ToArray();

                string strContent = "";
                foreach (string strLine in strValues)
                {
                    if (strLine.IndexOf("<add key=\"") < 0)
                    {
                        strContent += strLine + "\r\n";
                        continue;
                    }

                    //開始比對Key
                    int start = strLine.IndexOf("<add key=\"") + 10;
                    string Temp = strLine.Substring(start);
                    int end = Temp.IndexOf("\"");
                    if (strKey.Trim().ToUpper() == Temp.Substring(0, end).Trim().ToUpper())
                    {
                        int start2 = strLine.IndexOf("value=\"") + 7;

                        strContent += string.Format("{0}{1}{2}{3}",
                            strLine.Substring(0, start2),
                            Value.Trim(),
                            strLine.Substring(strLine.LastIndexOf("\"")),
                            "\r\n");
                    }
                    else
                        strContent += strLine + "\r\n";
                }

                File.WriteAllText(strPath, strContent);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        /// <summary>
        /// 取得檔案路徑
        /// </summary>
        /// <param name="strPath"></param>
        /// <param name="FileName"></param>
        /// <returns></returns>
        private string GetFilePath(string strPath = "", string FileName = "")
        {
            string ReturnPath = "";
            if (FileName.Trim() == "")
                FileName = "Setting.ini";
            if (strPath.Trim() != "")
            {
                if (strPath.IndexOf(".") > -1)//如果傳入的路徑已有檔名時
                    ReturnPath = strPath;
                else
                    ReturnPath = strPath + "\\" + FileName;
            }
            else//沒有傳入路徑時
            {
                if (striniFilePath != null && striniFilePath.Trim() != "")
                    ReturnPath = striniFilePath + "\\" + FileName;
                else
                    ReturnPath = System.AppDomain.CurrentDomain.BaseDirectory + "\\" + FileName;
            }

            return ReturnPath;
        }

        public string GetCSVString(string strPath, string strName)
        {
            string strfilePath = GetFilePath(strPath, strName);
            //string strCSV = System.IO.File.ReadAllText(strfilePath, System.Text.Encoding.GetEncoding("Big5"));//讀取CSV內容
            string strCSV = "";
            if (File.Exists(strfilePath))
                strCSV = System.IO.File.ReadAllText(strfilePath, System.Text.Encoding.UTF8);//讀取CSV內容

            return strCSV;
        }

        public DataTable GetCSVTable(string strName)
        {

            return GetCSVTable("", strName);
        }
        public DataTable GetCSVTable(string strPath, string strName)
        {
            string strfilePath = GetFilePath(strPath, strName);
            string strCSV = GetCSVString(strPath, strName);

            if (strCSV.Trim() == "") return null;
            string[] ContentRows = strCSV.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);//每一列資料
            if (ContentRows.Length < 1) return null;
            string[] Header = GetData(ContentRows[0]);//取出首列
            DataTable dt = new DataTable();

            //建立欄位
            for (int i = 0; i < Header.Length; i++)//每一個欄位
            {
                if (!dt.Columns.Contains(Header[i]))
                {
                    dt.Columns.Add(Header[i]);
                }
            }

            if (ContentRows.Length < 2) return dt;


            for (int i = 1; i < ContentRows.Length; i++)
            {
                try
                {
                    ContentRows[i] = ContentRows[i].Replace("\r", "");
                    if (ContentRows[i].Trim() == "") continue;
                    DataRow dr = dt.NewRow();
                    string[] Row = GetData(ContentRows[i]);//取出內容
                    for (int j = 0; j < Row.Length; j++)
                    {
                        dr[j] = Row[j].Trim();
                    }
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {

                    throw;
                }
            }


            return dt;
        }

        public DataTable GetCSVTable(string strPath, string strName, int ColumnCNT)
        {
            string strCSV = GetCSVString(strPath, strName);

            if (strCSV.Trim() == "") return null;
            string[] ContentRows = strCSV.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);//每一列資料
            if (ContentRows.Length < 1) return null;
            string[] Header = GetData(ContentRows[0]);//取出首列
            DataTable dt = new DataTable();

            //建立欄位
            for (int i = 0; i < Header.Length; i++)//每一個欄位
            {
                if (!dt.Columns.Contains(Header[i]))
                {
                    dt.Columns.Add(Header[i]);
                }
            }

            if (ContentRows.Length < 2) return dt;


            for (int i = 1; i < ContentRows.Length; i++)
            {
                try
                {
                    ContentRows[i] = ContentRows[i].Replace("\r", "");
                    if (ContentRows[i].Trim() == "") continue;
                    DataRow dr = dt.NewRow();
                    string[] Row = GetData(ContentRows[i]);//取出內容
                    for (int j = 0; j < ColumnCNT; j++)
                    {
                        dr[j] = Row[j].Trim();
                    }
                    dt.Rows.Add(dr);
                }
                catch (Exception ex)
                {

                    throw;
                }
            }


            return dt;
        }

        /// <summary>
        /// 將CSV整列字串截取各欄位資料
        /// </summary>
        /// <param name="strRow">依逗號隔開的字串</param>
        /// <returns>截取結果</returns>
        public string[] GetData(string strRow)//Joe研發新方式 CSV之列截取每個欄位資料
        {
            if (strRow.Trim() == "") return null;

            int Index1 = strRow.IndexOf("\"");
            int Index2 = strRow.IndexOf(",\"");
            strRow = strRow.Replace("\r", "");
            if (Index1 == -1)// 若內容沒有任何"
            {
                return strRow.Split(',');
            }


            //特殊處理
            List<string> listContents = new List<string>();

            bool IsStart = false;
            int indexStart = -1;
            int Count = 0;
            int intType = -1;//0:一般 1:有雙引號的資料

            //去除欄位開頭等號
            strRow = strRow.Replace(",=\"", ",\"");
            if (strRow.IndexOf("=\"") == 0)
                strRow = strRow.Substring(1);

            for (int i = 0; i < strRow.Length; i++)
            {
                if (strRow[i] == '\"' && intType != 0)//字元是雙引號 並且非一般模式時
                {
                    if (!IsStart)
                    {
                        IsStart = true;//已開始"資料
                        indexStart = i;
                        intType = 1;//雙引模式
                    }
                    else//已經有起始"時
                    {
                        if (i < strRow.Length - 1 && strRow[i + 1] != ',')//非最後並且 下一個字元不是,時
                        {
                            Count++;//累計
                        }

                        if (i < strRow.Length - 1 && strRow[i + 1] == ',')//非最後並且 下一個字元是,時
                        {
                            if (Count % 2 == 1)//表示還不是結束
                                Count++;
                            else//已經有偶數的"配對時 表示是結束的地方
                            {
                                //取得內容
                                string Temp = strRow.Substring(indexStart + 1, i - indexStart - 1);
                                listContents.Add(Temp.Replace("\"\"", "\""));
                                IsStart = false;
                                indexStart = -1;
                                Count = 0;
                                intType = -1;
                            }

                        }

                        if (i == strRow.Length - 1)//是最後時
                        {
                            //取得內容
                            string Temp = strRow.Substring(indexStart + 1, i - indexStart - 1);
                            listContents.Add(Temp.Replace("\"\"", "\""));
                            IsStart = false;
                            indexStart = -1;
                            Count = 0;
                            intType = -1;
                        }
                    }
                }
                else//字元非雙引號時
                {
                    if (intType != 1)//目前非雙引號模式時
                    {
                        if (!IsStart)//未開始時
                        {
                            if (strRow[i] == ',' && i < strRow.Length - 1 && strRow[i + 1] != '\"')
                            {
                                IsStart = true;
                                indexStart = i;
                                intType = 0;//一般模式
                            }
                        }
                        else//已經開始時
                        {
                            if (strRow[i] == ',')//結束一般模式
                            {
                                //取得內容
                                string Temp = strRow.Substring(indexStart + 1, i - indexStart - 1);
                                listContents.Add(Temp.Replace("\"\"", "\""));

                                IsStart = false;
                                indexStart = -1;
                                intType = -1;

                                if (i < strRow.Length - 1 && strRow[i + 1] != '\"')
                                {
                                    IsStart = true;
                                    indexStart = i;
                                    intType = 0;//一般模式
                                }
                            }
                        }
                    }
                }



            }


            //檢查補做一次

            if (IsStart)
            {
                //取得內容
                string Temp = strRow.Substring(indexStart + 1, strRow.Length - indexStart - 1);
                listContents.Add(Temp.Replace("\"\"", "\""));

                IsStart = false;
                indexStart = -1;
                intType = -1;
            }


            return listContents.ToArray();

        }

        public string SetCSVTable(DataTable dt, string strPath, string strName)
        {
            string strErrorMessage = "";
            try
            {
                string strfilePath = GetFilePath(strPath, strName);

                string strContents = "";
                string strHeader = "";
                foreach (DataColumn dc in dt.Columns)
                {
                    strHeader += string.Format("{0}\"{1}\"", strHeader.Trim() == "" ? "" : ",", dc.ColumnName.Trim()); ;
                }
                strContents += strHeader + "\r\n";
                    for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string strRow = "";
                    foreach (DataColumn dc in dt.Columns)
                    {
                        string Value = dt.Rows[i][dc.ColumnName].ToString();

                        Value = Value.Replace("\r", "").Replace("\n", "");
                        if (Value.IndexOf("\"") > -1 || Value.IndexOf(",") > -1)//若內容有出現"或, 要特別處理
                        {
                            Value = Value.Replace("\"", "\"\"");
                            strRow += string.Format("{0}\"{1}\"", strRow.Trim() == "" ? "" : ",", Value);
                        }
                        else//一般狀況
                            strRow += string.Format("{0}{1}", strRow.Trim() == "" ? "" : ",", Value);
                    }
                    strContents += strRow + "\r\n";

                }

                File.WriteAllText(strfilePath, strContents);
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message + ex.StackTrace;
            }
            return strErrorMessage;
        }

        public string SetCSVString(string strCSVContents, string strName)
        {
            return SetCSVString(strCSVContents, "", strName);
        }

        public string SetCSVString(string strCSVContents, string strPath, string strName)
        {
            string strErrorMessage = "";
            try
            {
                string strfilePath = GetFilePath(strPath, strName);
          
                File.WriteAllText(strfilePath, strCSVContents, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message + ex.StackTrace;
            }
            return strErrorMessage;
        }


    }

    //动态添加记录，ListView不闪烁
    class ListViewNF : System.Windows.Forms.ListView
    {
        public ListViewNF()
        {
            // Activate double buffering  
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);

            // Enable the OnNotifyMessage event so we get a chance to filter out   
            // Windows messages before they get to the form's WndProc  
            this.SetStyle(ControlStyles.EnableNotifyMessage, true);
        }

        protected override void OnNotifyMessage(Message m)
        {
            //Filter out the WM_ERASEBKGND message  
            if (m.Msg != 0x14)
            {
                base.OnNotifyMessage(m);
            }
        }
    }

    public class ListViewItemComparer : IComparer
    {
        public bool sort_b;
        public SortOrder order = SortOrder.Ascending;

        private int col;

        public ListViewItemComparer()
        {
            col = 0;
        }

        public ListViewItemComparer(int column, bool sort)
        {
            col = column;
            sort_b = sort;
        }

        public int Compare(object x, object y)
        {
            if (sort_b)
            {
                return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            }
            else
            {
                return String.Compare(((ListViewItem)y).SubItems[col].Text, ((ListViewItem)x).SubItems[col].Text);
            }
        }
    }

    //數字比較器
    public class ListViewItemComparerNum : IComparer
    {
        public bool sort_b;
        public SortOrder order = SortOrder.Ascending;

        private int col;

        public ListViewItemComparerNum()
        {
            col = 0;
        }

        public ListViewItemComparerNum(int column, bool sort)
        {
            col = column;
            sort_b = sort;
        }

        public int Compare(object x, object y)
        {
            decimal d1 = Convert.ToDecimal(((ListViewItem)x).SubItems[col].Text);
            decimal d2 = Convert.ToDecimal(((ListViewItem)y).SubItems[col].Text);
            if (sort_b)
            {
                return decimal.Compare(d1, d2);
            }
            else
            {
                return decimal.Compare(d2, d1);
            }
        }
    }
}

//public sealed class SiteHelper : Form
//{
//    public WebBrowser mBrowser = new WebBrowser();
//    ManualResetEvent mStart;
//    public event CompletedCallback Completed;
//    public SiteHelper(ManualResetEvent start)
//    {
//        mBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(mBrowser_DocumentCompleted);
//        mStart = start;
//    }
//    void mBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
//    {
//        // Generated completed event
//        Completed(mBrowser);
//    }
//    public void Navigate(string url)
//    {
//        // Start navigating
//        this.BeginInvoke(new Action(() => mBrowser.Navigate(url)));
//    }
//    public void Terminate()
//    {
//        // Shutdown form and message loop
//        this.Invoke(new Action(() => this.Close()));
//    }
//    protected override void SetVisibleCore(bool value)
//    {
//        if (!IsHandleCreated)
//        {
//            // First-time init, create handle and wait for message pump to run
//            this.CreateHandle();
//            this.BeginInvoke(new Action(() => mStart.Set()));
//        }
//        // Keep form hidden
//        value = false;
//        base.SetVisibleCore(value);
//    }
//}


//public abstract class SiteManager : IDisposable
//{
//    private ManualResetEvent mStart;
//    private SiteHelper mSyncProvider;
//    public event CompletedCallback Completed;

//    public SiteManager()
//    {
//        // Start the thread, wait for it to initialize
//        mStart = new ManualResetEvent(false);
//        Thread t = new Thread(startPump);
//        t.SetApartmentState(ApartmentState.STA);
//        t.IsBackground = true;
//        t.Start();
//        mStart.WaitOne();
//    }
//    public void Dispose()
//    {
//        // Shutdown message loop and thread
//        mSyncProvider.Terminate();
//    }
//    public void Navigate(string url)
//    {
//        // Start navigating to a URL
//        mSyncProvider.Navigate(url);
//    }
//    public void mSyncProvider_Completed(WebBrowser wb)
//    {
//        // Navigation completed, raise event
//        CompletedCallback handler = Completed;
//        if (handler != null)
//        {
//            handler(wb);
//        }
//    }
//    private void startPump()
//    {
//        // Start the message loop
//        mSyncProvider = new SiteHelper(mStart);
//        mSyncProvider.Completed += mSyncProvider_Completed;
//        Application.Run(mSyncProvider);
//    }
//}


//class Tester : SiteManager
//{
//    public Tester()
//    {
//        SiteEventArgs ar = new SiteEventArgs("MeSite");

//        base.Completed += new CompletedCallback(Tester_Completed);
//    }

//    void Tester_Completed(WebBrowser wb)
//    {
//        MessageBox.Show("Tester");
//        if (wb.DocumentTitle == "Hi")

//            base.mSyncProvider_Completed(wb);
//    }

//    //protected override void mSyncProvider_Completed(WebBrowser wb)
//    //{
//    //  //  MessageBox.Show("overload Tester");
//    //    //base.mSyncProvider_Completed(wb, ar);
//    //}
//}




