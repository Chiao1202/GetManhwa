using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace GetManhwa
{
    class Comicbus_Func
    {
        public DataTable GetAllBookDataTable(string strURL, out string Title)
        {
            Title = "";
            try
            {
                List<string> listBOOK = new List<string>();

                DataTable dt = new DataTable();
                dt.Columns.Add("SEL");//選取狀態
                dt.Columns.Add("Type");//漫畫站
                dt.Columns.Add("Name");//卷(話)名稱
                dt.Columns.Add("PageCount");//頁數
                dt.Columns.Add("URL");//卷(話)連結
                dt.Columns.Add("SubNum");//代碼
                dt.Columns.Add("IsDownLoad");//是否已下載
                dt.Columns.Add("Status");//狀態
                dt.Columns.Add("DateTime");//狀態

                List<Comicbus_Option> listComicbus_Option = new List<Comicbus_Option>();
                GetMenu(strURL, out Title, out listComicbus_Option);

                for (int i = 0; i < listComicbus_Option.Count; i++)
                {
                    Comicbus_Option comicbus_Option = listComicbus_Option[i];
                    if (comicbus_Option.Name.Trim() == "")
                        continue;
                    DataRow dr = dt.NewRow();
                    dr["SEL"] = true;
                    dr["Type"] = "comicbus";
                    dr["Name"] = comicbus_Option.Name;
                    dr["PageCount"] = (comicbus_Option.PageCount==null|| comicbus_Option.PageCount == 0)?-1:comicbus_Option.PageCount;
                    dr["URL"] = comicbus_Option.PageURL;
                    dr["SubNum"] = "";
                    dr["IsDownLoad"] = 0;
                    dr["Status"] = "未下載";
                    dt.Rows.Add(dr);
                }


                return dt;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public string GetMenu(string strURL, out string Title, out List<Comicbus_Option> listComicbus_Option)
        {
            listComicbus_Option = new List<Comicbus_Option>();
            Title = "";
            string strErrorMessage = "";
            WebClient wc = new WebClient();
            wc.Encoding = Encoding.GetEncoding(950);
            string Content = wc.DownloadString(strURL);

            //取得漫畫名稱
            Comm_Func.GetValue(Content, out Title, "<font color=\"#FFFFFF\" style=\"font-size:10pt; letter-spacing:1px\">", "</font");

            //遞迴取得所有目錄資料
            strErrorMessage = ForEachGetMenu(Content, ref listComicbus_Option);

            return strErrorMessage;
            for (int i = 0; i < listComicbus_Option.Count; i++)
            {
                Comicbus_Option comicbus_Option = listComicbus_Option[i];
                if (comicbus_Option.Name.Trim() == "")
                    continue;
                GetPageInfo(ref comicbus_Option,false,true);
            }

            //取得內頁取得總頁數 圖片



            return strErrorMessage;

        }
        //取得所有集數
        private string ForEachGetMenu(string content, ref List<Comicbus_Option> listComicbus_Option)
        {

            string strErrorMessage = "";
            if (content.IndexOf("<table id=\"rp_ct") > -1)
            {
                string NewContent = "";
                string Value = "";
                Comm_Func.GetValue(content, out NewContent, out Value, "<table id=\"rp_ct", "</table>");

                string[] TDs = Value.Split(new string[] { "<td" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < TDs.Length; i++)
                {
                    string strTD = TDs[i].Trim();
                    if (strTD.ToUpper().IndexOf("CVIEW") > -1)
                    {
                        try
                        {
                            string Temp02 = "";
                            Comm_Func.GetValue(strTD, out Temp02, "onclick=\"", ";");
                            string Temp03 = strTD.Replace("</td>", "").Replace("</a>", "").Replace("\r\n", "").Replace("<tr>", "").Replace("</tr>", "");
                            Temp03 = Temp03.Substring(Temp03.LastIndexOf(">") + 1);
                            Comicbus_Option Comicbus_Option = new Comicbus_Option();
                            Comicbus_Option.CVIEW = Temp02;
                            Comicbus_Option.Name = Temp03;

                            string Temp04 = "";
                            Comm_Func.GetValue(Temp02, out Temp04, "(", ")");
                            string[] para = Temp04.Split(new string[] { "," }, StringSplitOptions.None);
                            //function cview(url, catid, copyright)
                            string Temp05 = para[0].Replace("'", "");
                            Temp05 = Temp05.Replace(".html", "").Replace("-", ".html?ch=");
                            string ch = Temp05.Substring(Temp05.LastIndexOf("=") + 1);
                            //URL處理
                            string baseurl = "https://comicbus.live/";
                            if (para[2] == "1")
                                baseurl = "https://comicbus.live/online/c-";
                            else
                                baseurl = "https://comicbus.live/online/a-";
                            string PageURL = baseurl + Temp05;
                            Comicbus_Option.PageURL = PageURL;
                            Comicbus_Option.ch = ch;

                            var v = from Comicbus_Option t in listComicbus_Option
                                    where t.Name == Comicbus_Option.Name && t.PageURL == Comicbus_Option.PageURL
                                    select t;
                            if (v == null || v.Count() < 1)
                                listComicbus_Option.Add(Comicbus_Option);
                        }
                        catch (Exception ex)
                        {


                        }
                    }
                }


                ForEachGetMenu(NewContent, ref listComicbus_Option);
            }

            return strErrorMessage;
        }

        //目前不用一張一張取圖了
        public string GetPagePicture(string PageURL, out string PictureURL)
        {
            PictureURL = "";
            string strErrorMessage = "";
            try
            {
                Comicbus_Option comicbus_Option = new Comicbus_Option();
                comicbus_Option.PageURL = PageURL;//https://comicbus.live/online/a-18834.html?ch=1、https://comicbus.live/online/a-18834.html?ch=1-2
                GetPageInfo2(ref comicbus_Option,true, false);

                PictureURL = comicbus_Option.URL;
            }
            catch (Exception ex)
            {
                strErrorMessage = ex.Message + ex.StackTrace;
            }

            return strErrorMessage;
        }

        //取得總頁數跟所有圖片連結
        public void GetPageInfo(ref Comicbus_Option comicbus_Option, bool GetPageURL = true, bool GetPageCount = true)
        {
            string PageURL = comicbus_Option.PageURL;
            string ch = "1";//傳進來的ch
            int Index05 = PageURL.IndexOf("ch=");
            int Index06 = PageURL.IndexOf("&", Index05);
            if (Index05 > -1)
            {
                int Length = PageURL.Length - (Index05 + 3);
                if (Index06 > -1 && Index06 != Index05)
                    Length = Index06 - (Index05 + 3);
                ch = PageURL.Substring(Index05 + 3, Length);
            }
            WebClient wc = new WebClient();
            wc.Encoding = Encoding.GetEncoding(950);
            //內頁的圖片連結
            string Content2 = wc.DownloadString(PageURL); //"https://comicbus.live/online/a-18834.html?ch=1"

            //取得關鍵function
            string nviewJS = "";
            #region 從nview.js取相關圖片連結之function 及 頁數的function
            //string nview_JS = wc.DownloadString("https://comicbus.live/js/nview.js");//取得nview.js
            //string[] nview_JSs = nview_JS.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            //for (int i = 0; i < nview_JSs.Length; i++)
            //{
            //    //取得連結的function
            //    if (nview_JSs[i].IndexOf("function lc") > -1 || nview_JSs[i].IndexOf("function su") > -1
            //        || nview_JSs[i].IndexOf("function nn") > -1 || nview_JSs[i].IndexOf("function mm") > -1
            //        )
            //    {
            //        nviewJS += nview_JSs[i] + "\r\n";
            //    }
            //}

            nviewJS = @"
var y='46';
function spp(){};
var WWWWWTTTTTFFFFF='';
var document='';
function mm(p){return (parseInt((p-1)/10)%10)+(((p-1)%10)*3)};
function nn(n){return n<10?'00'+n:n<100?'0'+n:n;};
function su(a,b,c){var e=(a+'').substring(b,b+c);return (e);};
function lc(l){ 
if (l.length != 2) return l; 
var az = ""abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ""; 
var a = l.substring(0, 1); 
var b = l.substring(1, 2); 
if (a == ""Z"") 
return 8000 + az.indexOf(b); 
else 
return az.indexOf(a) * 52 + az.indexOf(b); 
};";
            #endregion

            Content2 = Content2.Replace("ge('TheImg').src=", "WWWWWTTTTTFFFFF=");
            //從內容頁取得關鍵連結資料
            string[] Scripts = Content2.Split(new string[] { "<script>" }, StringSplitOptions.RemoveEmptyEntries);
            #region 迴圈取得關鍵連結資料
            foreach (string item in Scripts)
            {
                if (item.IndexOf("function request") > -1)
                {
                    string Script_Org = item.Substring(0, item.IndexOf("</script>"));

                    //傳進來的ch
                    Script_Org = Script_Org.Replace("var ch=request('ch');", string.Format("var ch='{0}';", ch));
                    nviewJS += " " + Script_Org;
                    break;
                }
            }
            #endregion


            //取得頁數的function
            #region 取得頁數的function
            int PageCount = 0;
            string OnePageURL = "";
            List<string> listLink = new List<string>();
            if (GetPageCount && !GetPageURL)//如果要取得總頁數資料 不取連結 的話
            {
                nviewJS += "ps;";
                var PageTotal = Comm_Func.EvalJScript(nviewJS);// 
                if (!int.TryParse(PageTotal.ToString().Trim(), out PageCount))
                    PageCount = 0;
            }
            if (!GetPageCount && GetPageURL)//如果要不取總頁數 但取得連結資料的話
            {
                nviewJS += "\"https:\"+WWWWWTTTTTFFFFF;";
                var ResultURL = Comm_Func.EvalJScript(nviewJS);// 
                OnePageURL = ResultURL.ToString();
            }
            if (GetPageCount && GetPageURL)//如果要取得總頁數及取得連結資料的話
            {
                nviewJS += "ps+'__'+'https:'+WWWWWTTTTTFFFFF;";
                var Result = Comm_Func.EvalJScript(nviewJS);// 

                string[] strs = Result.ToString().Split(new string[] { "__" }, StringSplitOptions.RemoveEmptyEntries);
                OnePageURL = strs[1].ToString();
                if (!int.TryParse(strs[0].Trim(), out PageCount))
                    PageCount = 0;
                for (int i = 1; i <= PageCount; i++)
                {
                    string imgPath = "";
                    Comm_Func.GetValue(nviewJS, out imgPath, ";WWWWWTTTTTFFFFF='", ";");
                    imgPath = "'https:" + imgPath;
                    string _nviewJS = nviewJS + imgPath.Replace("(p)", "(" + i + ")") + ";";
                    listLink.Add(Comm_Func.EvalJScript(_nviewJS).ToString().Replace(strs[0].Trim()+"__", ""));
                }

            }
            #endregion


           
            if (GetPageCount)
                comicbus_Option.PageCount = PageCount;
            if (GetPageURL)
            {
                comicbus_Option.URL = OnePageURL;
                if (GetPageCount)
                    comicbus_Option.listPic = listLink;
            }

            #region 範例
            //            string RunJS2 = @"
            //var y=46;
            //function lc(l)
            //{
            //if(l.length!=2 )
            //return l;
            //var az='abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
            //var a=l.substring(0,1);
            //var b=l.substring(1,2);
            //if(a=='Z')
            //return 8000+az.indexOf(b);
            //else
            //return az.indexOf(a)*52+az.indexOf(b);
            //}
            //function su(a,b,c)
            //{
            //var e=(a+'').substring(b,b+c);
            //return (e);
            //}

            //function nn(n)
            //{
            //return n<10?'00'+n:n<100?'0'+n:n;
            //                                }
            //function mm(p){
            //   return (parseInt((p-1)/10)%10)+(((p-1)%10)*3)
            //};

            //var p=1;
            ////最後結果
            //var Result='';
            ////ch是傳進來的變數
            //var ch=1;
            //ch=parseInt(ch);
            //var pi=ch;
            //var ni=ch;
            //var ci=0;
            //var ps=0;
            //var chs=3;
            //var ti=18834;
            //var cs='bFdvcfBAXV6QQ74Q86BVAeS52pE8YyYPGPM99PHxuJabbebFp5tqM7R58UPdh4542T5mUCvc9bQrPMTSC7TMHspPacaMbFd396WSVPSP5xy3vWM4FtKR2hYjFwQNTEPtXHF8m2adaQ';
            //for(var i=0;i<3;i++)
            //{
            //    var gtnah= lc(su(cs,i*y+0,2));
            //    var wcbip= lc(su(cs,i*y+2,40));
            //    var opufb=lc(su(cs,i*y+42,2));
            //    var sowlt= lc(su(cs,i*y+44,2));
            //    ps=sowlt;
            //    if(opufb= ch){
            //        ci=i;
            //        Result='//img'+su(gtnah, 0, 1)+'.8comic.com/'+su(gtnah,1,1)+'/'+ti+'/'+opufb+'/'+ nn(p)+'_'+su(wcbip,mm(p),3)+'.jpg';
            //        break;
            //    }
            //}
            //"; 
            #endregion
        }

        //取得總頁數跟所有圖片連結
        public void GetPageInfo2(ref Comicbus_Option comicbus_Option, bool GetPageURL = true, bool GetPageCount = true)
        {
            string PageURL = comicbus_Option.PageURL;
            string ch = "1";//傳進來的ch
            int Index05 = PageURL.IndexOf("ch=");
            int Index06 = PageURL.IndexOf("&", Index05);
            if (Index05 > -1)
            {
                int Length = PageURL.Length - (Index05 + 3);
                if (Index06 > -1 && Index06 != Index05)
                    Length = Index06 - (Index05 + 3);
                ch = PageURL.Substring(Index05 + 3, Length);
            }
            WebClient wc = new WebClient();
            wc.Encoding = Encoding.GetEncoding(950);
            //內頁的圖片連結
            string Content2 = wc.DownloadString(PageURL); //"https://comicbus.live/online/a-18834.html?ch=1"

            //取得關鍵function
            string GetURLJS = "";
            string PageTotalJS = "";
            string nview_JS = wc.DownloadString("https://comicbus.live/js/nview.js");//取得nview.js
                                                                                     //從nview.js取相關圖片連結之function 及 頁數的function
            #region 從nview.js取相關圖片連結之function 及 頁數的function
            string[] nview_JSs = nview_JS.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < nview_JSs.Length; i++)
            {
                //取得連結的function
                if (nview_JSs[i].IndexOf("function lc") > -1 || nview_JSs[i].IndexOf("function su") > -1
                    || nview_JSs[i].IndexOf("function nn") > -1 || nview_JSs[i].IndexOf("function mm") > -1
                    )
                {
                    GetURLJS += nview_JSs[i] + "\r\n";
                }

                if (GetPageCount)//如果要取得總頁數資料的話
                {
                    //取得頁數的function
                    if (nview_JSs[i].IndexOf("function ss") > -1 || nview_JSs[i].IndexOf("function sp") > -1 || nview_JSs[i].IndexOf("var cs=") > -1)
                    {
                        PageTotalJS += nview_JSs[i] + "\r\n";
                    }
                }

            }
            #endregion




            //取得連結的function
            //Javascript內容處理
            GetURLJS = GetURLJS.Replace("var ch=request('ch');", string.Format("var ch='{0}';var Result = ''; ", ch));
            GetURLJS = GetURLJS.Replace("document.getElementById(e)", "null");




            //從內容頁取得關鍵連結資料
            string[] Scripts = Content2.Split(new string[] { "<script>" }, StringSplitOptions.RemoveEmptyEntries);
            #region 迴圈取得關鍵連結資料
            foreach (string item in Scripts)
            {
                if (item.IndexOf("function request") > -1)
                {
                    string Script_Org = item.Substring(0, item.IndexOf("</script>"));
                    string Temp01 = Script_Org.Substring(Script_Org.IndexOf("var ch=request('ch');"));

                    //傳進來的ch
                    Temp01 = Temp01.Replace("var ch=request('ch');", string.Format("ch='{0}';Result = ''; ", ch));
                    Temp01 = Temp01.Replace("document.getElementById(e)", "null");
                    Temp01 = Temp01.Replace("ge('TheImg').src=", "Result=");
                    Temp01 = Temp01.Replace("spp();", "");
                    GetURLJS += " " + Temp01;
                    break;
                }
            }
            #endregion

            #region 迴圈取得頁數資料
            if (GetPageCount)//如果要取得總頁數資料的話
            {
                foreach (string item in Scripts)
                {
                    if (item.IndexOf("var chs=") > -1)
                    {
                        string Script_Org = item.Substring(0, item.IndexOf("</script>"));
                        string Temp01 = Script_Org.Substring(Script_Org.IndexOf("var chs="));
                        Temp01 = Temp01.Substring(0, Temp01.IndexOf(";for") + 1);

                        PageTotalJS = Temp01 + " " + PageTotalJS;

                        //Temp01 = Temp01.Replace("var ch=request('ch');", string.Format("ch='{0}';Result = ''; ", ch));
                        //Temp01 = Temp01.Replace("document.getElementById(e)", "null");
                        //Temp01 = Temp01.Replace("ge('TheImg').src=", "Result=");
                        //Temp01 = Temp01.Replace("spp();", "");





                        break;
                    }
                }
            }
            #endregion

            //取得頁數的function
            #region 取得頁數的function
            int PageCount = 0;
            if (GetPageCount)//如果要取得總頁數資料的話
            {
                PageTotalJS = PageTotalJS.Substring(0, PageTotalJS.IndexOf("function jn("));
                string Temp05 = PageTotalJS.Replace("var ch=request('ch');", string.Format("var ch='{0}';", ch));
                Temp05 = "var TEMP01=null;function Temp(x,y,z){ TEMP01=1; }\r\n var nt = '';\r\n var pt = '';\r\n " + Temp05;
                Temp05 = Temp05.Replace("document.getElementById(", "Temp(");



                int Index01 = Temp05.IndexOf("function si");
                int Index02 = Temp05.IndexOf("}", Index01);
                string Temp06 = Temp05.Substring(0, Index01);
                string Temp07 = Temp05.Substring(Index02 + 1);
                Temp05 = Temp06 + Temp07;


                int Index03 = Temp05.IndexOf("function spp");
                int Index04 = Temp05.IndexOf("function sp(", Index03);
                string Temp08 = Temp05.Substring(0, Index03);
                string Temp09 = Temp05.Substring(Index04);
                Temp05 = Temp08 + Temp09;


                //int Index03 = Temp05.IndexOf("function jp");
                //int Index04 = Temp05.IndexOf("}", Index03);
                //string Temp08 = Temp05.Substring(0, Index03);
                //string Temp09 = Temp05.Substring(Index04 + 1);
                //Temp05 = Temp08 + Temp09;

                Temp05 = Temp05.Replace("initpage(", "Temp(");
                Temp05 = Temp05.Replace("si(", "Temp(");
                Temp05 = Temp05.Replace("alert(", "Temp(");
                Temp05 = Temp05.Replace("jump(", "Temp(");
                Temp05 = Temp05.Replace("jv(", "Temp(");
                Temp05 = Temp05.Replace("previd", "var previd");
                Temp05 = Temp05.Replace("nextid", "var nextid");
                Temp05 = Temp05.Replace("page =", "page=");
                Temp05 = Temp05.Replace("page=", "var page=");
                //Temp05 = Temp05.Replace("for (i", "for (var i");
                //Temp05 = Temp05.Replace("for(i", "for (var i");
                //Temp05 = Temp05.Replace("for( i", "for (var i");
                Temp05 = Temp05.Replace(".style.display = \"none\"", "");
                Temp05 = Temp05.Replace(".style.display = 'none;'", "");
                Temp05 = Temp05.Replace(".style.display = 'none'", "");
                Temp05 = Temp05.Replace(".style.display='none'", "");
                Temp05 = Temp05.Replace(".style.display='none;'", "");
                Temp05 = Temp05.Replace(".style.display='none;'", "");
                Temp05 = Temp05.Replace(".style.display=\"none\"", "");
                Temp05 = Temp05.Replace("Temp(\"lastchapter\")=ch;", "");
                Temp05 = Temp05.Replace("ge(\"lastchapter\")", "TEMP01");
                Temp05 = Temp05.Replace("ge('nextname')", "TEMP01");
                Temp05 = Temp05.Replace("ge('prevname')", "TEMP01");
                Temp05 = Temp05.Replace("ge('pagenum')", "TEMP01");
                Temp05 = Temp05.Replace("function ge(e){return Temp(e);}", "function ge(e){TEMP01=1;}");
                Temp05 = Temp05.Replace("XXXXXXXXXXXXX", "TEMP01");
                Temp05 = Temp05.Replace("XXXXXXXXXXXXX", "TEMP01");
                Temp05 = Temp05.Replace("XXXXXXXXXXXXX", "TEMP01");
                Temp05 = Temp05.Replace("XXXXXXXXXXXXX", "TEMP01");

                Temp05 = Temp05.Replace(".innerHTML", "");

                //方便閱讀
                Temp05 = Temp05.Replace("}", "}\r\n");
                //Temp05 = Temp05.Replace(";", ";\r\n");
                //Temp05 = Temp05.Replace(";\r\n'", ";'");

                string PageTotalJS_Run = string.Format("function PageTotal(){{ {0}  sp(); \r\n if(ps==null || ps=='') {{return 'NG';}} else {{return ps;}} }} PageTotal();", Temp05);//
                var PageTotal = Comm_Func.EvalJScript(PageTotalJS_Run);// 

                if (!int.TryParse(PageTotal.ToString().Trim(), out PageCount))
                    PageCount = 0;

            }
            #endregion


            GetURLJS = string.Format("function ResultURL() {{ {0} return Result; }} ResultURL();", GetURLJS);
            string URL = Comm_Func.EvalJScript(GetURLJS).ToString().Trim();//最後得到的連結
            //wc.DownloadFile("http:" + URL, "D:\\TEMP\\AAAAAAAAA.jpg");//http://img8.8comic.com/3/18834/1/001_dvc.jpg

            comicbus_Option.URL = URL;
            if (GetPageCount)
                comicbus_Option.PageCount = PageCount;

            #region 範例
            //            string RunJS2 = @"
            //var y=46;
            //function lc(l)
            //{
            //if(l.length!=2 )
            //return l;
            //var az='abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
            //var a=l.substring(0,1);
            //var b=l.substring(1,2);
            //if(a=='Z')
            //return 8000+az.indexOf(b);
            //else
            //return az.indexOf(a)*52+az.indexOf(b);
            //}
            //function su(a,b,c)
            //{
            //var e=(a+'').substring(b,b+c);
            //return (e);
            //}

            //function nn(n)
            //{
            //return n<10?'00'+n:n<100?'0'+n:n;
            //                                }
            //function mm(p){
            //   return (parseInt((p-1)/10)%10)+(((p-1)%10)*3)
            //};

            //var p=1;
            ////最後結果
            //var Result='';
            ////ch是傳進來的變數
            //var ch=1;
            //ch=parseInt(ch);
            //var pi=ch;
            //var ni=ch;
            //var ci=0;
            //var ps=0;
            //var chs=3;
            //var ti=18834;
            //var cs='bFdvcfBAXV6QQ74Q86BVAeS52pE8YyYPGPM99PHxuJabbebFp5tqM7R58UPdh4542T5mUCvc9bQrPMTSC7TMHspPacaMbFd396WSVPSP5xy3vWM4FtKR2hYjFwQNTEPtXHF8m2adaQ';
            //for(var i=0;i<3;i++)
            //{
            //    var gtnah= lc(su(cs,i*y+0,2));
            //    var wcbip= lc(su(cs,i*y+2,40));
            //    var opufb=lc(su(cs,i*y+42,2));
            //    var sowlt= lc(su(cs,i*y+44,2));
            //    ps=sowlt;
            //    if(opufb= ch){
            //        ci=i;
            //        Result='//img'+su(gtnah, 0, 1)+'.8comic.com/'+su(gtnah,1,1)+'/'+ti+'/'+opufb+'/'+ nn(p)+'_'+su(wcbip,mm(p),3)+'.jpg';
            //        break;
            //    }
            //}
            //"; 
            #endregion
        }
        //https://comicbus.live/online/a-1268.html?ch=1041
        //<table id = "rp_ctl04_0_dl_0" cellspacing="0" cellpadding="5" align="Center" style="width:95%;border-collapse:collapse;">
        //<table id = "rp_ctl05_0_dl_0" cellspacing="0" cellpadding="5" align="Center" style="width:95%;border-collapse:collapse;">
        //<table id = "rp_animelist_0_dl_0" cellspacing="0" cellpadding="5" align="Center" style="width:95%;border-collapse:collapse;">

        public void TEST(string PageURL)
        {
            string ch = "1";//傳進來的ch
            int Index05 = PageURL.IndexOf("ch=");
            int Index06 = PageURL.IndexOf("&", Index05);
            if (Index05 > -1)
            {
                int Length = PageURL.Length - (Index05 + 3);
                if (Index06 > -1 && Index06 != Index05)
                    Length = Index06 - (Index05 + 3);
                ch = PageURL.Substring(Index05 + 3, Length);
            }
            WebClient wc = new WebClient();
            wc.Encoding = Encoding.GetEncoding(950);
            //內頁的圖片連結
            string Content2 = wc.DownloadString(PageURL); //"https://comicbus.live/online/a-18834.html?ch=1"

            //取得關鍵function
            string PageTotalJS_TEST = "";
            string nview_JS = wc.DownloadString("https://comicbus.live/js/nview.js");//取得nview.js
                                                                                     //從nview.js取相關圖片連結之function 及 頁數的function
            #region 從nview.js取相關圖片連結之function 及 頁數的function
            string[] nview_JSs = nview_JS.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < nview_JSs.Length; i++)
            {

                /********測試****************/
                if (nview_JSs[i].IndexOf("function lc") > -1 || nview_JSs[i].IndexOf("function su") > -1
                   || nview_JSs[i].IndexOf("function nn") > -1 || nview_JSs[i].IndexOf("function mm") > -1
                   || nview_JSs[i].IndexOf("function ss") > -1 || nview_JSs[i].IndexOf("function sp") > -1
                   || nview_JSs[i].IndexOf("var cs=") > -1
                   )
                {
                    PageTotalJS_TEST += nview_JSs[i] + "\r\n";
                }

            }
            #endregion

            /********測試****************/
            PageTotalJS_TEST = PageTotalJS_TEST.Replace("var ch=request('ch');", string.Format("var ch='{0}';var Result = ''; ", ch));
            PageTotalJS_TEST = PageTotalJS_TEST.Replace("document.getElementById(e)", "null");


            PageTotalJS_TEST = @"var y='';
function spp(){};
var WWWWWTTTTTFFFFF='';
var document='';
function mm(p){return (parseInt((p-1)/10)%10)+(((p-1)%10)*3)};
function nn(n){return n<10?'00'+n:n<100?'0'+n:n;};
function su(a,b,c){var e=(a+'').substring(b,b+c);return (e);};
function lc(l){ 
if (l.length != 2) return l; 
var az = ""abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ""; 
var a = l.substring(0, 1); 
var b = l.substring(1, 2); 
if (a == ""Z"") 
return 8000 + az.indexOf(b); 
else 
return az.indexOf(a) * 52 + az.indexOf(b); 
};";

            //從內容頁取得關鍵連結資料
            string[] Scripts = Content2.Split(new string[] { "<script>" }, StringSplitOptions.RemoveEmptyEntries);

            #region 迴圈取得頁數資料

            foreach (string item in Scripts)
            {
                if (item.IndexOf("var chs=") > -1)
                {
                    /********測試****************/
                    string Script_Org2 = item.Substring(0, item.IndexOf("</script>"));

                    Script_Org2 = Script_Org2.Replace("ge('TheImg').src=", "WWWWWTTTTTFFFFF=");
                    PageTotalJS_TEST += " " + Script_Org2+" \r\n ps;";

                    //string Temp02 = Script_Org2.Substring(Script_Org2.IndexOf("var ch=request('ch');"));

                    ////傳進來的ch
                    //Temp02 = Temp02.Replace("var ch=request('ch');", string.Format("ch='{0}';Result = ''; ", ch));
                    //Temp02 = Temp02.Replace("document.getElementById(e)", "null");
                    //Temp02 = Temp02.Replace("ge('TheImg').src=", "Result=");
                    ////Temp02 = Temp02.Replace("spp();", "");
                    //PageTotalJS_TEST += " " + Temp02;

                    break;
                }
            }

            #endregion

            //取得頁數的function
            #region 取得頁數的function
            int PageCount = 0;
            /********測試****************/
            //string Temp10 = PageTotalJS_TEST.Replace("var ch=request('ch');", string.Format("var ch='{0}';", ch));
            //Temp10 = "var TEMP01=null;function Temp(x,y,z){ TEMP01=1; }\r\n var nt = '';\r\n var pt = '';\r\n function j(n){ TEMP01=1; }\r\n" + Temp10;
            //Temp10 = Temp10.Replace("document.getElementById(", "Temp(");
            //Temp10 = Temp10.Replace("initpage(", "Temp(");
            //Temp10 = Temp10.Replace("si(", "Temp(");
            //Temp10 = Temp10.Replace("alert(", "Temp(");
            //Temp10 = Temp10.Replace("jump(", "Temp(");
            //Temp10 = Temp10.Replace("jv(", "Temp(");
            //Temp10 = Temp10.Replace("previd", "var previd");
            //Temp10 = Temp10.Replace("nextid", "var nextid");
            //Temp10 = Temp10.Replace("page =", "page=");
            //Temp10 = Temp10.Replace("page=", "var page=");
            //Temp10 = Temp10.Replace(".style.display = \"none\"", "");
            //Temp10 = Temp10.Replace(".style.display = 'none;'", "");
            //Temp10 = Temp10.Replace(".style.display = 'none'", "");
            //Temp10 = Temp10.Replace(".style.display='none'", "");
            //Temp10 = Temp10.Replace(".style.display='none;'", "");
            //Temp10 = Temp10.Replace(".style.display='none;'", "");
            //Temp10 = Temp10.Replace(".style.display=\"none\"", "");
            //Temp10 = Temp10.Replace("Temp(\"lastchapter\")=ch;", "");
            //Temp10 = Temp10.Replace("ge(\"lastchapter\")", "TEMP01");
            //Temp10 = Temp10.Replace("ge('nextname')", "TEMP01");
            //Temp10 = Temp10.Replace("ge('prevname')", "TEMP01");
            //Temp10 = Temp10.Replace("ge('pagenum')", "TEMP01");
            //Temp10 = Temp10.Replace("function ge(e){return Temp(e);}", "function ge(e){TEMP01=1;}");
            //Temp10 = Temp10.Replace(".innerHTML", "");
            //方便閱讀
            //Temp10 = Temp10.Replace("}", "}\r\n");

            //string PageTotalJS_Run2 = string.Format("function PageTotal(){{ {0}  sp(); \r\n if(ps==null || ps=='') {{return 'NG';}} else {{return ps;}} }} PageTotal();", Temp10);//
            var PageTotal2 = Comm_Func.EvalJScript(PageTotalJS_TEST);//
            if (!int.TryParse(PageTotal2.ToString().Trim(), out PageCount))
                PageCount = 0;
            #endregion


        }
    }

    public class Comicbus_Option
    {
        public Comicbus_Option()
        {
        }

        public Comicbus_Option(string _PageURL)
        {
            PageURL = _PageURL;
        }

        public string Name { get; set; }
        public string CVIEW { get; set; }
        public string PageURL { get; set; }//內頁網址
        public string ch { get; set; }
        public string URL { get; set; }//最後圖片連結(第一頁)
        public int PageCount { get; set; }
        public List<string> listPic { get; set; }//所有內頁圖片連結
    }
}
