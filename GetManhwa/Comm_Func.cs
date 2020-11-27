using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetManhwa
{
    class Comm_Func
    {
        /// <summary>
        /// 取得關鍵字串後的所有內容
        /// </summary>
        /// <param name="Contents">原內容</param>
        /// <param name="NewContents">找到關鍵字之後下半部的內容</param>
        /// <param name="key">關鍵字串</param>
        public static void GetKey(string Contents, out string NewContents, string key)
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
        public static void GetValue(string Contents, out string Value, string keyS, string keyE)
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
        public static void GetValue(string Contents, out string NewContents, out string Value, string keyS, string keyE)
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

    }
}
