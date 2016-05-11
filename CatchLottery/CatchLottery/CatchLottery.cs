using Common.tools;
using EmailModule.Common.tools;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace tw.net.ebc {

    /// <summary>
    /// 彩券狀態回應列舉
    /// </summary>
    enum LotteryStatusCode {
        Success = 0,
        GetLotterySourceFail = 1,
        ProcessLotteryDataFail = 2,
        SaveLotteryDataFail = 3
    }

    /// <summary>
    /// 彩券代碼對應排序號類別
    /// </summary>
    class Lottery {
        public string code { get; set; }
        public short sort { get; set; }
    }

    class CatchLottery {

        #region *** 常數資料 ***

        //公佈最新開獎號碼小時區間
        private static readonly int ANNOUNCE_HOUR_BEGIN = 22;
        private static readonly int ANNOUNCE_HOUR_END = 23;

        //取得目前星期幾(若執行時間不是介於晚上公佈開獎結果之間，則取得前一日的星期)
        private static readonly DayOfWeek WEEK = ((DateTime.Now.Hour.Equals(ANNOUNCE_HOUR_BEGIN) || DateTime.Now.Hour.Equals(ANNOUNCE_HOUR_END)) ?
            DateTime.Now.DayOfWeek : DateTime.Now.AddDays(-1).DayOfWeek);

        //紀錄錯誤訊息資料夾
        private static readonly string LOG_FOLDER = string.Concat(Environment.CurrentDirectory, @"\App_Data\Log\");

        //彩券名稱分隔符號
        private static readonly string DELIMITER_INFO = Utility.getAppSettings("DELIMITER_INFO");

        //彩券開獎號碼分隔符號
        private static readonly string DELIMITER_NUMBER = Utility.getAppSettings("DELIMITER_NUMBER");

        //彩券名稱換行符號
        private static readonly string INFO_NEWLINE_SYMBO = "##INFO_NEWLINE_SYMBO##";

        //彩券開獎號碼換行符號
        private static readonly string NUMBER_NEWLINE_SYMBO = "##NUMBER_NEWLINE_SYMBO##";

        //彩券類型編號一(威力彩、大樂透、今彩539、38樂合彩、49樂合彩、39樂合彩、大福彩)
        private static readonly string[] LOTTERY_CODE_TYPE_1 = { "01", "02", "03", "07", "08", "09", "11" };

        //彩券類型編號二 (3星彩、4星彩)
        private static readonly string[] LOTTERY_CODE_TYPE_2 = { "05", "06" };

        //彩券類型有第二組號碼 (威力彩、大樂透)
        private static readonly string[] LOTTERY_SECOND_GROUP = { "01", "02" };

        //星期對應開獎類型編號對照表
        private static readonly Dictionary<DayOfWeek, string[]> WEEK_LOTTERY_DICY = new Dictionary<DayOfWeek, string[]> {
            {DayOfWeek.Monday, new string[] {"01", "07", "03", "09", "05", "06" } },
            {DayOfWeek.Tuesday, new string[] {"02", "08", "03", "09", "05", "06" } },
            {DayOfWeek.Wednesday, new string[] { "11", "03", "09", "05", "06" } },
            {DayOfWeek.Thursday, new string[] {"01", "07", "03", "09", "05", "06" } },
            {DayOfWeek.Friday, new string[] {"02", "08", "03", "09", "05", "06" } },
            {DayOfWeek.Saturday, new string[] { "11", "03", "09", "05", "06" } }
        };

        //彩券類型排序編號(參考 http://www.taiwanlottery.com.tw/runlottery/schedule.asp)
        private static readonly Dictionary<string, short> SORT_LOTTERY_DICY = new Dictionary<string, short> {
            { "01", 1}, { "07", 2}, { "02", 3}, { "08", 4}, { "11", 5}, { "03", 6}, { "09", 7}, { "05", 8}, { "06", 9}
        };

        //系統寄件者資訊
        private static readonly Dictionary<string, string> SYSTEM_MAIL_SENDER_DICY = new Dictionary<string, string>() {
            {"sender", "系統寄件者信箱"},
            {"account", "郵件伺服器帳號"},
            {"password", "郵件伺服器密碼"},
            {"smtp", "郵件伺服器"}
        };

        //系統郵件格式範本
        private static readonly string MAIL_TEMPLATE = File.ReadAllText(@"App_Data\Template\email-template.html", Encoding.UTF8);

        //主控台錯誤訊息
        private static readonly string ERROR_CONSOLE_MSG = "程式執行時發生錯誤！\n錯誤原因： {0}\n錯誤訊息： {1}\n\n請按任意鍵結束 ...";

        #endregion

        static void Main(string[] args) {

            Console.Write("--- 台灣彩券最新開獎資訊 ---\n\n");

            //優先檢查目前星期是否有進行開獎，而且文件檔是否已經是最新版本
            if (WEEK_LOTTERY_DICY.ContainsKey(WEEK) && (!checkIsLastVersion())) {
                HtmlDocument htmlDoc = new HtmlDocument();
                DateTime executeStartDt = DateTime.Now;
                DateTime executeEndDt;
                TimeSpan calcDiff;

                #region *** 抓取台灣彩券頁面最新開獎資訊 ***

                try {
                    Console.Write("正在擷取最新開獎資訊，此過程所需時間較長，請稍後 ...\n\n");
                    htmlDoc.LoadHtml(Utility.getWebContent(Utility.getAppSettings("LOTTERY_CATCH_PATH"), Encoding.UTF8));
                    executeEndDt = DateTime.Now;
                    calcDiff = executeEndDt.Subtract(executeStartDt);
                    Console.Write(string.Format("已完整取得最新開獎號碼，共花費： {0} 秒，開始產生文件 ...\n\n", Math.Round(calcDiff.TotalSeconds)));
                } catch (Exception ex) {
                    executeEndDt = DateTime.Now;
                    calcDiff = executeEndDt.Subtract(executeStartDt);
                    string exceptionTitle = "無法完整取得台灣彩券最新開獎號碼 (step.1)";
                    double totalSeconds = Math.Round(calcDiff.TotalSeconds);

                    //寄發錯誤通知郵件
                    string mailContent = string.Format(MAIL_TEMPLATE,
                        executeStartDt.ToString("yyyy-MM-dd HH:mm:ss"), exceptionTitle, totalSeconds, DateTime.Now.Year);
                    sendSystemEMail(mailContent);

                    //寫 Log 檔
                    writeSystemLog(exceptionTitle, ex, totalSeconds);

                    //顯示於主控台
                    string consoleMsg = string.Format(ERROR_CONSOLE_MSG, exceptionTitle, ex.Message);
                    Console.Write(consoleMsg);
                    Environment.Exit((int)LotteryStatusCode.GetLotterySourceFail);
                }

                #endregion

                #region *** 解析最新開獎資訊格式 ***

                List<string> lotteryContentList = new List<string>();
                try {

                    //存放彩券類別集合
                    List<Lottery> lotteryList = new List<Lottery>();

                    //主內容 html
                    HtmlNode contentHtml = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='right_full']");

                    //彩券編號標記集合
                    HtmlNodeCollection lotteryTypeTags = contentHtml.SelectNodes("//a[@name]");

                    //建立彩券類別並加入排序編號
                    foreach (var aTag in lotteryTypeTags) {
                        string key = aTag.Attributes["name"].Value;
                        if (SORT_LOTTERY_DICY.ContainsKey(key)) {
                            lotteryList.Add(new Lottery {
                                code = key,
                                sort = SORT_LOTTERY_DICY[key]
                            });
                        }
                    }

                    //將彩卷類別集合依照排序號碼由小到大排序後，組合相對應的樂透資訊字串
                    foreach (var lot in lotteryList.OrderBy(lot => lot.sort)) {
                        string typeCode = lot.code;        //彩券類型編號
                        short sort = lot.sort;      //彩券類型排序號
                        StringBuilder sb = new StringBuilder();

                        //判斷彩券類型編號是否屬於開獎項目
                        if (WEEK_LOTTERY_DICY[WEEK].Contains(typeCode)) {

                            //判斷彩券類型編號是否有在彩券類型編號陣列內
                            if (LOTTERY_CODE_TYPE_1.Contains(typeCode) || LOTTERY_CODE_TYPE_2.Contains(typeCode)) {

                                //取得開獎資訊並轉換指定格式
                                List<string> responseList = getLotteryInfo(typeCode, sort, contentHtml);

                                bool isChangedLine = false;      //判斷是否已出現換行符號並換行註記
                                foreach (var value in responseList) {

                                    //判斷值若為換行註記則加入換行符號
                                    if (!value.Equals(INFO_NEWLINE_SYMBO) && !value.Equals(NUMBER_NEWLINE_SYMBO)) {

                                        //分隔符號需判斷為彩券資訊或者開獎號碼而不同
                                        sb.Append(string.Concat(value, ((isChangedLine) ? DELIMITER_NUMBER : DELIMITER_INFO)));
                                    } else {
                                        sb.AppendLine();
                                        isChangedLine = true;
                                        continue;
                                    }
                                }

                                //儲存解析後的最新開獎資訊(移除最後一組分隔符號)
                                string sbString = sb.ToString();
                                lotteryContentList.Add(sbString.Remove(sbString.LastIndexOf(DELIMITER_NUMBER), DELIMITER_NUMBER.Length));
                            }
                        }
                    }
                } catch (Exception ex) {
                    executeEndDt = DateTime.Now;
                    calcDiff = executeEndDt.Subtract(executeStartDt);
                    string exceptionTitle = "解析最新開獎號碼發生未知的錯誤 (step.2)";
                    double totalSeconds = Math.Round(calcDiff.TotalSeconds);

                    //寄發錯誤通知郵件
                    string mailContent = string.Format(MAIL_TEMPLATE,
                        executeStartDt.ToString("yyyy-MM-dd HH:mm:ss"), exceptionTitle, totalSeconds, DateTime.Now.Year);
                    sendSystemEMail(mailContent);

                    //寫 Log 檔
                    writeSystemLog(exceptionTitle, ex, totalSeconds);

                    string consoleMsg = string.Format(ERROR_CONSOLE_MSG, exceptionTitle, ex.Message);
                    Console.Write(consoleMsg);
                    Environment.Exit((int)LotteryStatusCode.ProcessLotteryDataFail);
                }

                #endregion

                #region *** 最新開獎資訊儲存成實體檔 ***

                try {
                    //產生實體檔案
                    File.WriteAllLines(Utility.getAppSettings("SAVE_FILE_PATH"), lotteryContentList, Encoding.UTF8);
                } catch (Exception ex) {
                    executeEndDt = DateTime.Now;
                    calcDiff = executeEndDt.Subtract(executeStartDt);
                    string exceptionTitle = "產生最新開獎號碼文件發生未知的錯誤 (step.3)";
                    double totalSeconds = Math.Round(calcDiff.TotalSeconds);

                    //寄發錯誤通知郵件
                    string mailContent = string.Format(MAIL_TEMPLATE,
                        executeStartDt.ToString("yyyy-MM-dd HH:mm:ss"), exceptionTitle, totalSeconds, DateTime.Now.Year);
                    sendSystemEMail(mailContent);

                    //寫 Log 檔
                    writeSystemLog(exceptionTitle, ex, totalSeconds);

                    string consoleMsg = string.Format(ERROR_CONSOLE_MSG, exceptionTitle, ex.Message);
                    Console.Write(consoleMsg);
                    Environment.Exit((int)LotteryStatusCode.SaveLotteryDataFail);
                }

                #endregion

                Console.Write(string.Format("程序執行成功！最新開獎號碼文件存放位置： {0}", Utility.getAppSettings("SAVE_FILE_PATH")));
            }
        }

        #region *** 邏輯函式 ***

        /// <summary>
        /// 檢查目前已存在的文件檔是否為最新版本
        /// </summary>
        /// <returns>bool</returns>
        private static bool checkIsLastVersion() {
            bool result = false;

            if (Utility.getAppSettings("IS_CHECK_LASTVERSION").Equals("1")) {
                string filePath = Utility.getAppSettings("SAVE_FILE_PATH");

                if (File.Exists(filePath)) {
                    result = true;
                    string[] content = File.ReadAllLines(filePath);

                    //若執行時間不是介於晚上公佈開獎結果之間，則取得前一日的日期
                    DateTime catchDate = ((DateTime.Now.Hour.Equals(ANNOUNCE_HOUR_BEGIN) || DateTime.Now.Hour.Equals(ANNOUNCE_HOUR_END)) ?
                        DateTime.Now : DateTime.Now.AddDays(-1));
                    string catchDateFormat = string.Format("{0}年{1}月{2}日", (catchDate.Year - 1911), catchDate.Month, catchDate.Day);

                    //判斷讀取到的彩券開獎日期是否與目前的日期相同，若不相同則表示文件檔不是最新的
                    foreach (string data in content) {

                        //只比較有彩券名稱分隔符號的列資料
                        if (data.Contains(DELIMITER_INFO)) {

                            //取得每一種彩券的開獎日期
                            string lotteryDate = data.Split(new string[] { DELIMITER_INFO }, StringSplitOptions.RemoveEmptyEntries)[1];
                            if (!lotteryDate.Equals(catchDateFormat)) {
                                result = false;
                                break;
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 取得樂透中獎號碼資訊
        /// </summary>
        /// <param name="typeCode">彩券編號</param>
        /// <param name="sort">彩券序號</param>
        /// <param name="contentHtml">html 內容</param>
        /// <returns>List<string></returns>
        private static List<string> getLotteryInfo(string typeCode, short sort, HtmlNode contentHtml) {
            List<string> resultList = new List<string>();

            //樂透中獎號碼資訊元素
            HtmlNodeCollection winningNumberRows = contentHtml.SelectNodes(string.Format("//a[@name='{0}']/following-sibling::div[1]//table[1]/tr", typeCode));

            if (winningNumberRows != null) {
                resultList.Add(winningNumberRows[0].SelectSingleNode("./td[2]").InnerText.Trim());      //遊戲名稱
                resultList.Add(winningNumberRows[2].SelectSingleNode("./td[2]/span[1]").InnerText.Trim().Replace(" ", string.Empty));      //開獎日期
                resultList.Add(winningNumberRows[1].SelectSingleNode("./td[2]/span[1]").InnerText.Trim());      //期別
                resultList.Add(INFO_NEWLINE_SYMBO);
                resultList.Add(sort.ToString());

                //中獎號碼
                if (LOTTERY_CODE_TYPE_1.Contains(typeCode)) {        //中獎號碼有分大小順序與開出順序(威力彩、大樂透、今彩539、38樂合彩、49樂合彩、39樂合彩、大福彩)

                    //取得所有獎號有文字節點的span標籤
                    HtmlNodeCollection spanTags = winningNumberRows[4]
                        .SelectNodes("./td[2]/br/preceding-sibling::span/descendant-or-self::span/child::text()");

                    //取得span標籤內符合開獎號碼規則的獎號
                    foreach (HtmlNode span in spanTags) {
                        string spanValue = span.InnerHtml.Trim().Replace("&nbsp;", string.Empty);
                        if (!string.IsNullOrEmpty(spanValue)) {
                            resultList.Add(spanValue);
                        }
                    }

                } else {        //中獎號碼無分大小順序與開出順序(3星彩、4星彩)
                    winningNumberRows[4]
                        .SelectSingleNode("./td[2]/span[1]")
                        .InnerText
                        .Trim()
                        .Replace("&nbsp;", string.Empty)
                        .ToCharArray()
                        .ToList()
                        .ForEach(num => resultList.Add(num.ToString()));
                }

                resultList.Add(NUMBER_NEWLINE_SYMBO);
            }

            return resultList;
        }

        /// <summary>
        /// 寄發系統通知信
        /// </summary>
        /// <param name="mailContent">信件內容</param>
        private static void sendSystemEMail(string mailContent) {
            if (Utility.getAppSettings("IS_SEND_MAIL").Equals("1")) {
                EmailMod email = new EmailMod(SYSTEM_MAIL_SENDER_DICY["smtp"], SYSTEM_MAIL_SENDER_DICY["account"], SYSTEM_MAIL_SENDER_DICY["password"], false);

                email.SendHtmlMail(Utility.getAppSettings("MAIL_SEND_TO"),      //收件者
                                                  SYSTEM_MAIL_SENDER_DICY["sender"],     //寄件者郵件
                                                  Utility.getAppSettings("MAIL_SENDER_NAME"),       //寄件者名稱
                                                  Utility.getAppSettings("MAIL_SUBJECT"),       //主旨
                                                  mailContent,      //郵件內容
                                                  false,        //是否使用密件副本寄送
                                                  Convert.ToInt32(Utility.getAppSettings("MAIL_SLEEP_TIME")));      //若寄信失敗間隔多久再次寄送
            }
        }

        /// <summary>
        /// 建立系統紀錄文字檔
        /// </summary>
        /// <param name="exceptionTitle">例外標題</param>
        /// <param name="ex">Exception 例外物件</param>
        /// <param name="seconds">花費秒數</param>
        private static void writeSystemLog(string exceptionTitle, Exception ex, double seconds) {
            if (Utility.getAppSettings("IS_WRITE_LOG").Equals("1")) {

                //Log 內容
                string content = string.Format(@"錯誤標題： {0}{1}錯誤訊息： {2}{1}花費時間： {3} 秒{1}錯誤位置： {4}:{5}",
                    exceptionTitle, Environment.NewLine, ex.Message, seconds, Utility.getExLineNumber(ex), Utility.getExColumnNumber(ex));

                //建立 log 檔案
                Utility.createTextFile(LOG_FOLDER, string.Format("{0}.txt", DateTime.Now.ToString("yyyyMMddHHmmss")), content);
            }
        }

        #endregion
    }
}
