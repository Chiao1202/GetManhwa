using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;

namespace GetSetting
{

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
                string strPath = GetFilePath();//取得檔案路徑
                FileInfo F = new FileInfo(strPath);
                string[] strValues = File.ReadLines(F.FullName).ToArray();
                foreach (string strLine in strValues)
                {
                    if (strLine.IndexOf("<add key=\"") < 0) continue;
                    int start = strLine.IndexOf("<add key=\"")+10;
                    string Temp = strLine.Substring(start);
                    int end = Temp.IndexOf("\"");
                    if (strKey.Trim().ToUpper() == Temp.Substring(0, end).Trim().ToUpper())
                    {
                        int start2 = Temp.IndexOf("value=\"")+7;
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

         public string GetSetValue(string strPath,string strKey)
         {
             try
             {
                string strValue = "";
                strPath = GetFilePath(strPath, "Setting.ini");
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

        public bool SetValue(string strPath,string strKey, string Value)
        {

            try
            {
                if(strPath.Trim()=="")
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

        public string GetCSVString(string strPath,string strName)
        {
            string strfilePath = GetFilePath(strPath, strName);
            //string strCSV = System.IO.File.ReadAllText(strfilePath, System.Text.Encoding.GetEncoding("Big5"));//讀取CSV內容
            string strCSV = System.IO.File.ReadAllText(strfilePath, System.Text.Encoding.UTF8);//讀取CSV內容

            return strCSV;
        }

        public DataTable GetCSVTable(string strPath, string strName)
        {
            string strCSV = GetCSVString(strPath,strName);

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

        public DataTable GetCSVTable(string strPath, string strName,int ColumnCNT)
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
                    if (!IsStart && i == 0)
                    {
                        IsStart = true;//已開始"資料
                        indexStart = -1;
                        intType = 0;//雙引模式
                    }
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

        public string SetCSVTable(DataTable dt, string strPath,string strName)
        {
            string strErrorMessage = "";
            try
            {
                string strfilePath = GetFilePath(strPath, strName);

                string strContents = "";
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

        public string SetCSVString(string strCSVContents, string strPath, string strName)
        {
            string strErrorMessage = "";
            try
            {
                string strfilePath = GetFilePath(strPath, strName);

                File.WriteAllText(strfilePath, strCSVContents,Encoding.UTF8);
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message + ex.StackTrace;
            }
            return strErrorMessage;
        }
    }
}
