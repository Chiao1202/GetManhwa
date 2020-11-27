using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace GetManhwa
{
    class Manhuagui_Func
    {
        public Manhuagui_Func()
        {

        }

        public string GetManhuaguiURL(string SouceURL,out List<string>  listURL)
        {
            listURL = new List<string>();
            string strErrorMessage = "";
            int Index = SouceURL.IndexOf("#");
            if(Index>-1)
                SouceURL = SouceURL.Substring(0, Index);
            try
            {
                //判斷SouceURL格式是否為網址
                if (!IsUrl(SouceURL))
                {
                    strErrorMessage = "來源網址非正常URL";
                    goto END;
                }

                //取得網頁內容
                byte[] WebContentByte = GetWebContent(SouceURL);
                //轉成字串
                string WebContent = System.Text.Encoding.UTF8.GetString(WebContentByte);

                //取得關鍵JavaScript
                string Value01 = "";
                GetValue(WebContent, out Value01, "(function(", "{}))");
                //再將function外層包回來
                string orin_scripts = "(function(" + Value01 + "{}))";

                //取得function中抓出最後一組的Base64，為了進行解碼
                string[] m = orin_scripts.Split(new string[] { ",'" }, StringSplitOptions.None);
                string _sc = "";
                GetValue(m[m.Count() - 1], out _sc, "", "'[");

                //解碼_sc
                string _c = DecompressFromBase64(_sc);
                //將解碼完之內容放回原JavaScript function
                string new_scripts = orin_scripts.Replace(_sc, _c);

                //把 【['\x73\x70\x6c\x69\x63']('\x7c')】，換成【.split('|')】
                new_scripts = new_scripts.Replace(@"['\x73\x70\x6c\x69\x63']('\x7c')", ".split('|')");
                new_scripts += ";";


                //執行JavaScript跑出結果
                string finalScripts = Comm_Func.EvalJScript(new_scripts).ToString();
                //截取結果的內容
                string Value_Result = "";
                GetValue(finalScripts, out Value_Result, "SMH.imgData(", ").preInit();");

                //取得path
                string path = "";
                GetValue(Value_Result, out path, "\"path\":\"", "\"");
                string path_encode = HttpUtility.UrlEncode(path).Replace("%2f", "/");

                var ms = new MemoryStream(Encoding.Unicode.GetBytes(Value_Result));
                DataContractJsonSerializer deseralizer = new DataContractJsonSerializer(typeof(Manhuagui_Option));
                Manhuagui_Option manhuagui_Option = (Manhuagui_Option)deseralizer.ReadObject(ms);
                ms.Close();
                ms.Dispose();
                for (int i = 0; i < manhuagui_Option.files.Count; i++)
                {
                    listURL .Add(string.Format("https://i.hamreus.com{0}{1}?e={2}&m={3}\r\n",
                                          manhuagui_Option.path_encode.Trim(),
                                          manhuagui_Option.files[i].Trim(),
                                          manhuagui_Option.sl.e.ToString().Trim(),
                                          manhuagui_Option.sl.m.Trim()
                      ));
                }


            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message + ex.StackTrace;
            }

            END:

            return strErrorMessage;
        }


        //判斷字串是否為網址
        public static bool IsUrl(string str)
        {
            try
            {
                string Url = @"^http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?$";
                return Regex.IsMatch(str, Url);
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public byte[] GetWebContent(string url)
        {
            try
            {
                if (url.ToLower().IndexOf("http:") > -1 || url.ToLower().IndexOf("https:") > -1)
                {
                    // URL                 

                    HttpWebRequest request = null;
                    HttpWebResponse response = null;
                    byte[] byteData = null;

                    request = (HttpWebRequest)WebRequest.Create(url);
                    request.Timeout = 60000;
                    request.Proxy = null;
                    request.UserAgent = "user_agent','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.75 Safari/537.36";
                    //request.Referer = getSystemKey("HTTP_REFERER");
                    response = (HttpWebResponse)request.GetResponse();
                    Stream stream = response.GetResponseStream();
                    byteData = ReadStream(stream, 32765);
                    response.Close();
                    stream.Close();
                    return byteData;
                }
                else
                {
                    /*System.IO.StreamReader sr = new System.IO.StreamReader(url);

                    string sContents = sr.ReadToEnd();
                    sr.Close();
                    return s2b(sContents);
                    */
                    /*FileStream fs = new FileStream(url, FileMode.Open);
                    byte[] buffer = new byte[fs.Length];
                    fs.Read(buffer, 0, buffer.Length);
                    fs.Close();
                    return buffer;
                    */
                    byte[] data;
                    using (StreamReader sr = new StreamReader(url))
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            sr.BaseStream.CopyTo(ms);
                            data = ms.ToArray();
                            ms.Close();
                            ms.Dispose();
                        }
                        sr.Close();
                        sr.Dispose();
                    };
                    return data;
                }
            }
            catch (Exception ex)
            {

                throw;
            }
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


        #region 解壓縮Base64解碼
        private const string KeyStrBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
        private static readonly IDictionary<char, char> KeyStrBase64Dict = CreateBaseDict(KeyStrBase64);
        private static IDictionary<char, char> CreateBaseDict(string alphabet)
        {
            var dict = new Dictionary<char, char>();
            for (var i = 0; i < alphabet.Length; i++)
            {
                dict[alphabet[i]] = (char)i;
            }
            return dict;
        }
        public static string DecompressFromBase64(string input)
        {
            if (input == null) throw new ArgumentNullException(nameof(input));

            return Decompress(input.Length, 32, index => KeyStrBase64Dict[input[index]]);
        }

        private static string Decompress(int length, int resetValue, Func<int, char> getNextValue)
        {
            var dictionary = new List<string>();
            var enlargeIn = 4;
            var numBits = 3;
            string entry;
            var result = new StringBuilder();
            int i;
            string w;
            int bits = 0, resb, maxpower, power;
            var c = '\0';

            var data_val = getNextValue(0);
            var data_position = resetValue;
            var data_index = 1;

            for (i = 0; i < 3; i += 1)
            {
                dictionary.Add(((char)i).ToString());
            }

            maxpower = (int)Math.Pow(2, 2);
            power = 1;
            while (power != maxpower)
            {
                resb = data_val & data_position;
                data_position >>= 1;
                if (data_position == 0)
                {
                    data_position = resetValue;
                    data_val = getNextValue(data_index++);
                }
                bits |= (resb > 0 ? 1 : 0) * power;
                power <<= 1;
            }

            switch (bits)
            {
                case 0:
                    bits = 0;
                    maxpower = (int)Math.Pow(2, 8);
                    power = 1;
                    while (power != maxpower)
                    {
                        resb = data_val & data_position;
                        data_position >>= 1;
                        if (data_position == 0)
                        {
                            data_position = resetValue;
                            data_val = getNextValue(data_index++);
                        }
                        bits |= (resb > 0 ? 1 : 0) * power;
                        power <<= 1;
                    }
                    c = (char)bits;
                    break;
                case 1:
                    bits = 0;
                    maxpower = (int)Math.Pow(2, 16);
                    power = 1;
                    while (power != maxpower)
                    {
                        resb = data_val & data_position;
                        data_position >>= 1;
                        if (data_position == 0)
                        {
                            data_position = resetValue;
                            data_val = getNextValue(data_index++);
                        }
                        bits |= (resb > 0 ? 1 : 0) * power;
                        power <<= 1;
                    }
                    c = (char)bits;
                    break;
                case 2:
                    return "";
            }
            w = c.ToString();
            dictionary.Add(w);
            result.Append(c);
            while (true)
            {
                if (data_index > length)
                {
                    return "";
                }

                bits = 0;
                maxpower = (int)Math.Pow(2, numBits);
                power = 1;
                while (power != maxpower)
                {
                    resb = data_val & data_position;
                    data_position >>= 1;
                    if (data_position == 0)
                    {
                        data_position = resetValue;
                        data_val = getNextValue(data_index++);
                    }
                    bits |= (resb > 0 ? 1 : 0) * power;
                    power <<= 1;
                }

                int c2;
                switch (c2 = bits)
                {
                    case (char)0:
                        bits = 0;
                        maxpower = (int)Math.Pow(2, 8);
                        power = 1;
                        while (power != maxpower)
                        {
                            resb = data_val & data_position;
                            data_position >>= 1;
                            if (data_position == 0)
                            {
                                data_position = resetValue;
                                data_val = getNextValue(data_index++);
                            }
                            bits |= (resb > 0 ? 1 : 0) * power;
                            power <<= 1;
                        }

                        c2 = dictionary.Count;
                        dictionary.Add(((char)bits).ToString());
                        enlargeIn--;
                        break;
                    case (char)1:
                        bits = 0;
                        maxpower = (int)Math.Pow(2, 16);
                        power = 1;
                        while (power != maxpower)
                        {
                            resb = data_val & data_position;
                            data_position >>= 1;
                            if (data_position == 0)
                            {
                                data_position = resetValue;
                                data_val = getNextValue(data_index++);
                            }
                            bits |= (resb > 0 ? 1 : 0) * power;
                            power <<= 1;
                        }
                        c2 = dictionary.Count;
                        dictionary.Add(((char)bits).ToString());
                        enlargeIn--;
                        break;
                    case (char)2:
                        return result.ToString();
                }

                if (enlargeIn == 0)
                {
                    enlargeIn = (int)Math.Pow(2, numBits);
                    numBits++;
                }

                if (dictionary.Count - 1 >= c2)
                {
                    entry = dictionary[c2];
                }
                else
                {
                    if (c2 == dictionary.Count)
                    {
                        entry = w + w[0];
                    }
                    else
                    {
                        return null;
                    }
                }
                result.Append(entry);

                // Add w+entry[0] to the dictionary.
                dictionary.Add(w + entry[0]);
                enlargeIn--;

                w = entry;

                if (enlargeIn == 0)
                {
                    enlargeIn = (int)Math.Pow(2, numBits);
                    numBits++;
                }
            }
        }

        #endregion

        //執行JavaScript跑出結果
        public static Microsoft.JScript.Vsa.VsaEngine Engine = Microsoft.JScript.Vsa.VsaEngine.CreateEngine();
        public static object EvalJScript(string JScript)
        {
            object Result = null;
            try
            {
                Result = Microsoft.JScript.Eval.JScriptEvaluate(JScript, Engine);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return Result;
        }


        /*****已有的********/
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
    }


    public class Manhuagui_Option
    {
        public int bid { get; set; }
        public string bname { get; set; }
        public string bpic { get; set; }
        public int cid { get; set; }
        public string cname { get; set; }
        public List<string> files { get; set; }
        public string finished { get; set; }
        public int len { get; set; }
        public string path { get; set; }

        string _path_encode = "";
        public string path_encode
        {
            get
            {
                if (path != null && path.Trim() != "")
                    return HttpUtility.UrlEncode(path).Replace("%2f", "/");
                else
                    return "";
            }
            set { _path_encode = value; }
        }

        public int status { get; set; }
        public string block_cc { get; set; }
        public int nextId { get; set; }
        public int prevId { get; set; }
        public sl sl { get; set; }
    }
    public class sl
    {
        public int e { get; set; }

        public string m { get; set; }

    }
}

