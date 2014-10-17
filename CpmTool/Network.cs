using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Drawing;
using System.Net;
using System.Collections;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.DirectoryServices.Protocols;
using System.ServiceModel.Security;
using System.Text.RegularExpressions;

namespace CpmTool
{
    public delegate void DelegateDidGetData(List<List<string>> data, bool isSuccess, int dbIndex);

    class Network
    {
        private CookieContainer CC = new CookieContainer();
        private static readonly string DefaultUserAgent = "mozilla/5.0 (windows nt 5.1) applewebkit/537.11 (khtml, like gecko) "
                                                + "chrome/23.0.1271.95 safari/537.11";
        private static readonly int RequestTimeout = 30000;
        private int siteType;
        private int requireType;
        private string username;
        private string password;
        private int dbIndex;
        private int timezone = 0;
        
        public event DelegateDidGetData onDidGetData;

        public Network(int sitetype, int dbIndex, int requireType, string user, string pwd, DelegateDidGetData del, int timezone)
        {
            this.onDidGetData += new DelegateDidGetData(del);
            this.requireType = requireType;
            this.siteType = sitetype;
            this.username = user;
            this.password = pwd;
            this.dbIndex = dbIndex;
            this.timezone = timezone;
        }

        public void run()
        {
            switch(this.siteType)
            {
                case 0: run_yieldmanager(); break;
                case 1: run_onlinemedia(); break;
            }
        }

        private String DoGet(String url, String referer)
        {
            String data = "";
            HttpWebRequest webReqst = (HttpWebRequest)WebRequest.Create(url);
            webReqst.Method = "GET";
            webReqst.UserAgent = DefaultUserAgent;
            webReqst.KeepAlive = true;
            webReqst.CookieContainer = CC;
            webReqst.Referer = referer;
            webReqst.Timeout = RequestTimeout;

            try
            {
                HttpWebResponse webResponse = (HttpWebResponse)webReqst.GetResponse();
                /*foreach (Cookie c in webResponse.Cookies)
                {
                    c.Domain = c.Domain.Replace(".www", "www");
                }
                CC.Add(webResponse.Cookies);*/
                BugFix_CookieDomain(CC);

                if (webResponse.StatusCode == HttpStatusCode.OK && webResponse.ContentLength < 1024 * 1024)
                {
                    StreamReader reader = new StreamReader(webResponse.GetResponseStream(), Encoding.Default);
                    data = reader.ReadToEnd();
                }
            }
            catch(Exception ex)
            {
                Console.Write("http Exception ### " + ex.Message);
            }

            return data;
        }

        private String DoPost(String url, String Content, String referer)
        {
            string html = "";
            HttpWebRequest webReqst = null;
            //如果是发送HTTPS请求  
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                webReqst = WebRequest.Create(url) as HttpWebRequest;
                webReqst.ProtocolVersion = HttpVersion.Version10;
            }
            else
            {
                webReqst = WebRequest.Create(url) as HttpWebRequest;
            }

            webReqst.Method = "POST";
            webReqst.UserAgent = DefaultUserAgent;
            webReqst.ContentType = "application/x-www-form-urlencoded";
            webReqst.ContentLength = Content.Length;
            webReqst.CookieContainer = CC;
            webReqst.Referer = referer;
            webReqst.Timeout = RequestTimeout;

            try
            {
                byte[] data = Encoding.Default.GetBytes(Content);
                Stream stream = webReqst.GetRequestStream();
                stream.Write(data, 0, data.Length);


                HttpWebResponse webResponse = (HttpWebResponse)webReqst.GetResponse();
                /*foreach (Cookie c in webResponse.Cookies)
                {
                    c.Domain.Replace(".www", "www");    //修复CookieContainer的bug
                }
                CC.Add(webResponse.Cookies);*/
                BugFix_CookieDomain(CC);

                if (webResponse.StatusCode == HttpStatusCode.OK && webResponse.ContentLength < 1024 * 1024)
                {
                    StreamReader reader = new StreamReader(webResponse.GetResponseStream(), Encoding.Default);
                    html = reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Console.Write("http Exception ### " + ex.Message);
            }

            return html;
        }

        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true; //总是接受  
        }

        private void BugFix_CookieDomain(CookieContainer cookieContainer)
        {
            System.Type _ContainerType = typeof(CookieContainer);
            Hashtable table = (Hashtable)_ContainerType.InvokeMember("m_domainTable",
                                       System.Reflection.BindingFlags.NonPublic |
                                       System.Reflection.BindingFlags.GetField |
                                       System.Reflection.BindingFlags.Instance,
                                       null,
                                       cookieContainer,
                                       new object[] { });
            ArrayList keys = new ArrayList(table.Keys);
            foreach (string keyObj in keys)
            {
                string key = (keyObj as string);
                if (key[0] == '.')
                {
                    string newKey = key.Remove(0, 1);
                    table[newKey] = table[keyObj];
                }
            }
        }

#region yieldmanager works


        private string reportUrl(string rand, int type)
        {
            string end_hour = "";
            string start_hour = "";
            string quick_date = "";
            string interval = "";
            switch (type/10)
            {
            case 0:     //last 24 hours
                end_hour = "23";
                start_hour = "0";
                quick_date = "last24";
            	break;
            case 1:     //last month
                end_hour = "23";
                start_hour = "0";
                quick_date = "lastmonth";
                break;
            case 2:     //month to date
                end_hour = "7";
                start_hour = "3";
                quick_date = "mtd";
                break;
            }
            switch (type%10)
            {
            case 0:     //none
                interval = "none";
            	break;
            case 1:     //day
                interval = "day";
                break;
            }

            string url = "https://my.yieldmanager.com/tab.php?";
            url += "from_report_page=1&report_ready=0&savesettings=0&total_max_rows=0&submit_report.x=19&submit_report.y=12"
                 + "&quick_date=" + quick_date
                 + "&interval=" + interval
                 + "&timezone=1&start_date=" + Uri.EscapeDataString(DateTime.Now.ToString("MM/dd/yyyy"))
                 + "&start_hour=" + start_hour 
                 + "&end_date=" + Uri.EscapeDataString(DateTime.Now.ToString("MM/dd/yyyy"))
                 + "&end_hour=" + end_hour
                 + "&metricsOption=default"
                 + "&filtering_io_id=on&filtering_line_item_id=on&filtering_site_id=on&filtering_section_id=on&filtering_size_id=on"
                 + "&filtering_pop_type_id=on&filtering_country_woe_id=on&filtering_country_group_id=on&filtering_age_gender=on"
                 + "&filtering_frequency=on&filtering_screenType=on&filtering_mobileWifi=on&inc=11&tab_id=1"
                 + "&rand=" + rand + "#report_section";
            return url;
        }

        private string reportUrl2(string rand, string param, int type)
        {

            string end_hour = "";
            string start_hour = "";
            string quick_date = "";
            string interval = "";
            switch (type / 10)
            {
                case 0:     //last 24 hours
                    end_hour = "23";
                    start_hour = "0";
                    quick_date = "last24";
                    break;
                case 1:     //last month
                    end_hour = "23";
                    start_hour = "0";
                    quick_date = "lastmonth";
                    break;
                case 2:     //month to date
                    end_hour = "7";
                    start_hour = "3";
                    quick_date = "mtd";
                    break;
            }
            switch (type % 10)
            {
                case 0:     //none
                    interval = "none";
                    break;
                case 1:     //day
                    interval = "day";
                    break;
            }

            string url = "https://my.yieldmanager.com/tab.php?";
            url += "from_report_page=1&savesettings=0&total_max_rows=0&submit_report.x=35&submit_report.y=12&"
                 + "quick_date=" + quick_date
                 + "&interval=" + interval
                 + "&timezone=1"
                 + "&start_date=" + DateTime.Now.ToString("MM/dd/yyyy")
                 + "&start_hour=" + start_hour
                 + "&end_date=" + DateTime.Now.ToString("MM/dd/yyyy")
                 + "&end_hour=" + end_hour
                 + "&metricsOption=default&filtering_io_id=on&filtering_line_item_id=on&filtering_site_id=on"
                 + "&filtering_section_id=on&filtering_size_id=on&filtering_pop_type_id=on&filtering_country_woe_id=on"
                 + "&filtering_country_group_id=on&filtering_age_gender=on&filtering_frequency=on"
                 + "&filtering_screenType=on&filtering_mobileWifi=on&inc=11&tab_id=1"
                 + "&rand=" + rand
                 + "&report_ready=1&report_url=" + param
                 + "&tstamp=" + GetTimeLikeJS();
            return url;
        }

        private void run_yieldmanager()
        {
            const string loginStr = "https://my.yieldmanager.com/index.php";
            const string logoutStr = "https://my.yieldmanager.com/index.php?logout=1";
            const string reportStr = "https://my.yieldmanager.com/tab.php?tab_id=1&inc=11";
            const string waitStr = "https://my.yieldmanager.com/reports/report_ajax.php?make_requests=y";
            const string checkStr = "https://my.yieldmanager.com/reports/report_ajax.php?check_tokens=y&t[]=";

            const string Interval2DayUrl = "https://my.yieldmanager.com/reports/changeGrouping.php?action=change_grouping&type=publisher&filter=Interval&value=1&intervalStr=day&inc=11&art=";
            const string Interval2NoneUrl = "https://my.yieldmanager.com/reports/changeGrouping.php?action=change_grouping&type=publisher&filter=Interval&value=0&intervalStr=none&inc=11&art=";

        Login:
            string content = "redir=&username="
                            + this.username
                            + "&password="
                            + this.password
                            + "&x=32&y=14";
            string html = DoPost(loginStr, content, "");
            if (html == "" || html.IndexOf("logout=1") == -1)
            {
                //登录失败
                goto Login;
            }

        GetData:

            string getReportURL = reportStr;
            html = DoGet(getReportURL, "");
            if (html == "" || html.IndexOf("rand") == -1)
            {
                goto GetData;
            }

            //requireType   个位表示Interval    十位表示Date range
            //00:last 24 hours / none
            //10:last month    / none
            //20:month to date / none
            //21:month to date / day
            string IntervalUrl = "";
            if (this.requireType % 10 == 0)
            {
                //none
                IntervalUrl = Interval2NoneUrl;
            }
            else
            {
                //day
                IntervalUrl = Interval2DayUrl;
            }

            DoGet(IntervalUrl, getReportURL);

            string rand = GetMid(html, "name=\"rand\" value=\"", "\"");
            string strReportUrl = reportUrl(rand, this.requireType);
            string referUrl = strReportUrl;
            if (DoGet(strReportUrl, getReportURL) == "")
            {
                goto GetData;
            }
            html = DoGet(waitStr, referUrl);
            if (html == "")
            {
                goto GetData;
            }
            string tmp = GetMid(html, "[\"", "\"]");
            tmp = tmp.Replace("\\", "");
            strReportUrl = checkStr + Uri.EscapeDataString(tmp);
            html = DoGet(strReportUrl, referUrl);
            if (html == "")
            {
                goto GetData;
            }
            tmp = GetMid(html, "[\"", "\"]");
            tmp = tmp.Replace("\\", "");
            strReportUrl = reportUrl2(rand, tmp, this.requireType);
            html = DoGet(strReportUrl, referUrl);
            if (html != "")
            {
                //处理数据
                if (html.IndexOf("No data returned") != -1)
                {
                    this.onDidGetData(null, false, this.dbIndex);
                }
                else
                {
                    string tableHtml = html.Remove(0, html.IndexOf("<table xmlns"));
                    List<List<string>> list = getHtmlTableData(tableHtml);
                    this.onDidGetData(list, true, this.dbIndex);
                }
                
            }
        }
#endregion

#region onlinemedia works


        private string getIdPostContent(int type)
        {
            string range = "";
            string interval = "";
            switch (type / 10)
            {
                case 0:     //last 24 hours
                    range = "last_48_hours";
                    break;
                case 1:     //last month
                    range = "last_month";
                    break;
                case 2:     //month to date
                    range = "month_to_date";
                    break;
                case 3:
                    range = "yesterday" + "&report%5Bgroup_by%5D%5B%5D=placement_id";
                    break;
            }
            switch (type % 10)
            {
                case 0:     //none
                    interval = "cumulative";
                    break;
                case 1:     //day
                    interval = "day" + "report%5Bfixed_columns%5D%5B%5D=day&report%5Bgroup_by%5D%5B%5D=day";
                    break;
            }
            //report[group_by][]:placement_id
            string timezone = "Asia%2FHong_Kong";
            switch (this.timezone)
            {
            case 0:
                timezone = "Asia%2FHong_Kong";
            	break;
            case 1:
                timezone = "EST5EDT";
                break;
            case 2:
                timezone = "UTC";
                break;
            default:
                timezone = "Asia%2FHong_Kong";
                break;
            }
            //Asia%2FHong_Kong
            //UTC
            //EST5EDT
            string content  = "report%5Bcategory%5D=publisher_login&report%5Btype%5D=analytics&report%5Bformat%5D=standard"
                            + "&report%5Brange%5D=" + range
                            + "&report%5Bstart_date%5D=&report%5Bend_date%5D="
                            + "&report%5Binterval%5D=" + interval
                            + "&report%5Btimezone%5D=" + timezone
                            + "&report%5Bmetrics%5D%5B%5D=imps_total"
                            + "&report%5Bmetrics%5D%5B%5D=clicks&report%5Bmetrics%5D%5B%5D=total_convs&report%5Bmetrics%5D%5B%5D=publisher_revenue"
                            + "&report%5Bmetrics%5D%5B%5D=publisher_rpm&report%5Bshow_usd_currency%5D=true&report%5Brun_type%5D=run_now"
                            + "&report%5Bemail_format%5D=excel&report%5Bpre_send_now_email_addresses%5D=miguel%40creafi-online-media.com"
                            + "&report%5Bschedule_when%5D=daily"
                            + "&report%5Bschedule_format%5D=excel&report%5Bschedule_email_addresses%5D=&report%5Bname%5D=&report%5Btimezone%5D="
                            + timezone;
            return content;
        }


        private string getStatusPostContent(string id, string pageid)
        {
            double timeRand = GetTimeLikeJS();
            double back = new Random().NextDouble() * 100;
            timeRand += (int)back;
            string request_id = timeRand.ToString() + "." + back.ToString().Split('.')[1];
            string report_id = "report_id%5B" + id + "%5D=" + id;
            string page_id = "page_id=" + pageid;
            string content = request_id + "&" + report_id + "&" + page_id;
            int length = content.Length;
            length += 13;
            int old_length_num_chars = length.ToString().Length;
            length += old_length_num_chars;
            int new_length_num_chars = length.ToString().Length;
            length += (new_length_num_chars - old_length_num_chars);
            content = "body_length=" + length.ToString() + "&" + content;
            return content;
        }

        private string getPostContent(string ready_id, string pageid)
        {
            string isDayStr = "";
            if (this.requireType %10 == 1)
            {
                //day
                isDayStr = "&columns%5B%5D=day";
            }
            switch (this.requireType / 10)
            {
                case 3:
                    //yesterday
                    isDayStr = "&columns%5B%5D=placement";
                    break;
            }

            double timeRand = GetTimeLikeJS();
            double back = new Random().NextDouble() * 100;
            timeRand += (int)back;
            string request_id = timeRand.ToString() + "." + back.ToString().Split('.')[1];
            string content = "request_id=" + request_id
                           + "&id=" + ready_id
                           + isDayStr
                           + "&columns%5B%5D=imps_total&columns%5B%5D=clicks&columns%5B%5D=total_convs&columns%5B%5D=publisher_revenue"
                           + "&columns%5B%5D=publisher_rpm&show_as_pivot=false&report_type=publisher_analytics"
                           + "&page_id=" + pageid;
            int length = content.Length;
            length += 13;
            int old_length_num_chars = length.ToString().Length;
            length += old_length_num_chars;
            int new_length_num_chars = length.ToString().Length;
            length += (new_length_num_chars - old_length_num_chars);
            content = "body_length=" + length.ToString() + "&" + content;
            return content;
        }

        private List<List<String>> processDatas(string html)
        {
            List<List<String>> list = new List<List<String>>();
            //处理头部
            string data = GetMid(html, "\"header\":\"", "\",");
            data = data.Replace("\\", "");
            List<string> headList = new List<string>();
            while(data.IndexOf("<th") != -1)
            {
                headList.Add(ExecRepaceHTML(getInnerhtml(data, "th")));
                data = data.Remove(0, data.IndexOf("</th>") + 5);
            }
            list.Add(headList);
            data = GetMid(html, "\"html\":\"", "\",");
            data = data.Replace("\\", "");
            while (data.IndexOf("<tr") >= 0)
            {
                string tmpTd = getInnerhtml(data, "tr");
                List<string> strList = new List<string>();
                while (tmpTd.IndexOf("<td") >= 0)
                {
                    strList.Add(ExecRepaceHTML(getInnerhtml(tmpTd, "td")));
                    tmpTd = tmpTd.Remove(0, tmpTd.IndexOf("</td>") + 5);
                }
                if (strList.Count > 0)
                {
                    list.Add(strList);
                }
                data = data.Remove(0, data.IndexOf("</tr>") + 5);
            }
            if (this.requireType / 10 == 3)
            {
                //yesterday with id
                data = GetMid(html, "\"total\":\"", "\",");
                data = data.Replace("\\", "");
                while (data.IndexOf("<tr") >= 0)
                {
                    string tmpTd = getInnerhtml(data, "tr");
                    List<string> strList = new List<string>();
                    while (tmpTd.IndexOf("<td") >= 0)
                    {
                        strList.Add(ExecRepaceHTML(getInnerhtml(tmpTd, "td")));
                        tmpTd = tmpTd.Remove(0, tmpTd.IndexOf("</td>") + 5);
                    }
                    if (strList.Count > 0)
                    {
                        list.Add(strList);
                    }
                    data = data.Remove(0, data.IndexOf("</tr>") + 5);
                }
            }
            return list;
        }

        private void run_onlinemedia()
        {
            const string loginUrl = "https://console.appnexus.com/index/sign-in";
            string getIdUrl = "https://console.appnexus.com/{0}/report/get-id";
            string getStatusUrl = "https://console.appnexus.com/{0}/report/check-report-status";
            string getUrl = "https://console.appnexus.com/{0}/report/get";

        Login:
            string html = "";
            string content = "redir=&app_id=&app_redirect=&username=" + this.username
                           + "&password=" + this.password;
            html = DoPost(loginUrl, content, "");
            if (html == "" || html.IndexOf("Sign Out") == -1)
            {
                goto Login;
            }

        GetData:

            string page_id = GetMid(html, "page_id = '", "'");
            //增加ui_version格式化url  2013/11/27 
            string ui_version = GetMid(html, "ui_version = '", "'");
            getIdUrl = string.Format(getIdUrl, ui_version);
            getStatusUrl = string.Format(getStatusUrl, ui_version);
            getUrl = string.Format(getUrl, ui_version);

            //string test = "report%5Bcategory%5D=publisher_login&report%5Btype%5D=analytics&report%5Bformat%5D=standard&report%5Brange%5D=yesterday&report%5Bstart_date%5D=&report%5Bend_date%5D=&report%5Binterval%5D=cumulative&report%5Btimezone%5D=EST5EDT&report%5Bmetrics%5D%5B%5D=imps_total&report%5Bmetrics%5D%5B%5D=clicks&report%5Bmetrics%5D%5B%5D=total_convs&report%5Bmetrics%5D%5B%5D=publisher_revenue&report%5Bmetrics%5D%5B%5D=publisher_rpm&report%5Bshow_usd_currency%5D=true&report%5Bgroup_by%5D%5B%5D=placement_id&report%5Brun_type%5D=run_now&report%5Bemail_format%5D=excel&report%5Bpre_send_now_email_addresses%5D=lpita%40sonital.com&report%5Bschedule_when%5D=daily&report%5Bschedule_format%5D=excel&report%5Bschedule_email_addresses%5D=&report%5Bname%5D=&report%5Btimezone%5D=EST5EDT";
            html = DoPost(getIdUrl, getIdPostContent(this.requireType), "");
            //html = DoPost(getIdUrl, test, "");
            if (html == "" || html.IndexOf("report_id") == -1)
            {
                goto GetData;
            }
            string reportId = GetMid(html, "\"report_id\":\"", "\"");
            html = DoPost(getStatusUrl, getStatusPostContent(reportId, page_id), "");
            if (html == "" || html.IndexOf("OK") == -1)
            {
                goto GetData;
            }
            string readyId = GetMid(html,"\"ready\":[\"", "\"");
            html = DoPost(getUrl, getPostContent(readyId, page_id), "");
            if (html == "" || html.IndexOf("OK") == -1)
            {
                goto GetData;
            }
            List<List<string>> result = processDatas(html);
            if (result.Count<2)
            {
                this.onDidGetData(result, false, this.dbIndex);
            }
            else
            {
                if (Form1.isForWG)
                {
                    List<List<string>> tmp = new List<List<string>>();
                    result[0].RemoveAt(0);//删除placement字段，因为会显示total而已
                    result[result.Count - 1].RemoveAt(0);//删除placement字段，因为会显示(total/唯一的ID)而已
                    tmp.Add(result[0]);
                    tmp.Add(result[result.Count - 1]);
                    result = tmp;
                    this.onDidGetData(result, true, this.dbIndex);
                    return;
                }

                if (this.requireType / 10 == 3)
                {
                    //yesterday with id
                    List<List<string>> tmp = new List<List<string>>();
                    result[0].Add("ID数");
                    result[0].Add("平均流量");
                    tmp.Add(result[0]);
                    if (result.Count > 2)
                    {
                        //ID数
                        int imps = -1;
                        for (int i = 0; i < result[0].Count; i++ )
                        {
                            if (result[0][i].Contains("Imps"))
                            {
                                imps = i;
                                break;
                            }
                        }
                        int id_sum = result.Count - 2;
                        if (imps != -1)
                        {
                            //去掉imps<10的ID
                            for (int i = 1; i < result.Count - 1; i++)
                            {
                                int imp = int.Parse(result[i][imps].Replace(",", ""));
                                if (imp <= 10)
                                {
                                    id_sum--;
                                }
                            }
                        }
                        result[result.Count - 1].Add(id_sum.ToString());
                        int temp = int.Parse(result[result.Count - 1][1].Replace(",", ""));
                        if (id_sum != 0)
                        {
                            temp = temp / id_sum / 24;
                        }
                        else
                        {
                            temp = 0;
                        }
                        result[result.Count - 1].Add(temp.ToString());
                    }
                    else if (result.Count == 2)
                    {
                        //如果只有2个，表示该号只有一个ID
                        result[result.Count - 1].Add("1");
                        int temp = int.Parse(result[result.Count - 1][1].Replace(",", ""));
                        temp = temp / 24;
                        result[result.Count - 1].Add(temp.ToString());
                    }
                    else
                    {
                        result[result.Count - 1].Add("");
                        result[result.Count - 1].Add("");
                    }
                    result[0].RemoveAt(0);//删除placement字段，因为会显示total而已
                    result[result.Count - 1].RemoveAt(0);//删除placement字段，因为会显示(total/唯一的ID)而已
                    tmp.Add(result[result.Count - 1]);
                    result = tmp;
                }
                this.onDidGetData(result, true, this.dbIndex);
            }
        }


#endregion
        


#region 功能函数
        private List<List<string>> getHtmlTableData(string src)
        {
            int index = src.IndexOf("<table");
            string tmp = getInnerhtml(src, "table");

            if (tmp == "")
            {
                return null;
            }

            List<List<string>> list = new List<List<string>>();
            if(tmp.IndexOf("<thead") >= 0)
            {
                //获取表头
                string tmpTr = getInnerhtml(tmp, "thead");
                while(tmpTr.IndexOf("<tr") >= 0)
                {
                    string tmpTh = getInnerhtml(tmpTr, "tr");
                    List<string> strList = new List<string>();
                    while (tmpTh.IndexOf("<th") >= 0)
                    {
                        strList.Add(ExecRepaceHTML(getInnerhtml(tmpTh, "th")));
                        tmpTh = tmpTh.Remove(0, tmpTh.IndexOf("</th>") + 5);
                    }
                    if (strList.Count>0)
                    {
                        list.Add(strList);
                    }
                    tmpTr = tmpTr.Remove(0, tmpTr.IndexOf("</tr>") + 5);
                }
                tmp = tmp.Remove(0, tmp.IndexOf("</thead>") + 8);
            }
            if (tmp.IndexOf("<tbody") >= 0)
            {
                //获取表内容
                string tmpTr = getInnerhtml(tmp, "tbody");
                while (tmpTr.IndexOf("<tr") >= 0)
                {
                    string tmpTd = getInnerhtml(tmpTr, "tr");
                    List<string> strList = new List<string>();
                    while (tmpTd.IndexOf("<td") >= 0)
                    {
                        strList.Add(ExecRepaceHTML(getInnerhtml(tmpTd, "td")));
                        tmpTd = tmpTd.Remove(0, tmpTd.IndexOf("</td>") + 5);
                    }
                    if (strList.Count > 0)
                    {
                        list.Add(strList);
                    }
                    tmpTr = tmpTr.Remove(0, tmpTr.IndexOf("</tr>") + 5);
                }
            }
            
            return list;
        }


        private string ExecRepaceHTML(string Htmlstring)   
        {   
            //去除HTML标签 
            Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
            return Htmlstring;  
        }  

        private string getInnerhtml(string src, string meta)
        {
            //获取innerHTML，例如<a>bbb</a>，则获取bbb
            int index = src.IndexOf("<" + meta);
            string tmp = "";
            if (index != -1)
            {
                tmp = src.Remove(0, index);
                tmp = GetMid(tmp, ">", "</" + meta);
            }
            return tmp;
        }

        private long GetTimeLikeJS()
        {
            long lLeft = 621355968000000000;
            DateTime dt = DateTime.Now;
            long Sticks = (dt.Ticks - lLeft) / 10000;
            return Sticks;
        }

        private String GetMid(String input, String s, String e)
        {
            int pos = input.IndexOf(s);
            if (pos == -1)
            {
                return "";
            }

            pos += s.Length;

            int pos_end = 0;
            if (e == "")
            {
                pos_end = input.Length;
            }
            else
            {
                pos_end = input.IndexOf(e, pos);
            }

            if (pos_end == -1)
            {
                return "";
            }

            return input.Substring(pos, pos_end - pos);
        }

        
#endregion
        
    }
}
