using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Text.RegularExpressions;
using System.Linq;
using System.Globalization;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string monthsString = "5"; // 月份字串
            string datesString = "16"; // 日期字串

            string[] months = monthsString.Split(','); // 分割月份字串為陣列
            string[] dates = datesString.Split(','); // 分割日期字串為陣列

            // 初始化 ChromeDriver
            IWebDriver driver = new ChromeDriver();

            // 建立 DatabaseReader
            string connectionString = "Server=192.168.10.230;Database=ZDB;User Id=sa;Password=#HQTP@9090;";
            var dbReader = new DatabaseReader(connectionString);
            var factoryData = dbReader.GetFactoryData();

            Thread.Sleep(5000);

            foreach (var data in factoryData)
            {
                string factoryid = data.Item1;
                string username = data.Item2;
                string password = data.Item3;

                Console.WriteLine($"FactoryID: {factoryid}, Username: {username}, Password: {password}");

                // 访问目标网页
                driver.Navigate().GoToUrl("https://hvcs.taipower.com.tw/Account/NewLogon");
                Thread.Sleep(5000);

                // 输入用户名和密码
                IWebElement usernameInput = FindElement(driver, By.Id("UserName"), 10);
                usernameInput.Clear(); // 清除輸入框
                usernameInput.SendKeys(username);

                IWebElement passwordInput = FindElement(driver, By.Id("Password"), 10);
                passwordInput.Clear(); // 清除輸入框
                passwordInput.SendKeys(password);

                // 点击登录按钮
                IWebElement loginButton = FindElement(driver, By.ClassName("btn_Sign_in"), 10);
                loginButton.Click();
                Thread.Sleep(5000);

                try
                {
                    // 等待对话框出现并取消
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(AlertIsPresent());
                    IAlert alert = driver.SwitchTo().Alert();
                    alert.Dismiss();
                }
                catch (TimeoutException)
                {
                    Console.WriteLine("对话框未出现或未在预期的时间内出现。");
                }

                Thread.Sleep(5000);

                // 调用 click_element 方法，模拟点击操作
                if (!ClickElement(driver, "//*[@id='sb-site']/div[6]/div[2]/div/ul/li[1]/a"))
                    continue;

                Thread.Sleep(5000);

                if (factoryid == "LD-T1HIST")
                {
                    // 调用 click_element 方法，模拟点击操作
                    if (!ClickElement(driver, "//*[@id='Clickbox']/div/input[1]"))
                    {
                        continue;
                    }
                }

                Thread.Sleep(5000);
                driver.Navigate().GoToUrl("https://hvcs.taipower.com.tw/Customer/Module/PowerAnalyze");

                for (int i = 0; i < months.Length; i++)
                {
                    int month = int.Parse(months[i]);
                    int date = int.Parse(dates[i]);

                    // 建立 SQL 指令碼
                    string delete_sql = $@"DELETE FROM Test_TPC
                                WHERE FactoryID = '{factoryid}' 
                                AND Dtime >= DATETIMEFROMPARTS(YEAR(getdate()), {month}, {date}, 0, 15, 0, 0)
                                AND DTime <= DATEADD(DAY, 1, DATEFROMPARTS(YEAR(getdate()), {month}, {date}))";

                    // 建立連線物件
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        // 開啟連線
                        connection.Open();

                        // 建立命令物件
                        using (SqlCommand command = new SqlCommand(delete_sql, connection))
                        {
                            // 執行 SQL 指令碼
                            command.ExecuteNonQuery();
                        }
                    }
                    Console.WriteLine("SQL 指令碼執行完成。");

                    // 獲取當前年份
                    int current_year = DateTime.Now.Year;

                    // 如果月份和日期是單個數字，則使用0進行填充
                    string monthPadded = month.ToString().PadLeft(2, '0');
                    string dayPadded = date.ToString().PadLeft(2, '0');

                    // 組合成日期字符串
                    string dateStr = $"{current_year}-{monthPadded}-{dayPadded}";

                    // 將日期字符串轉換為 DateTime 對象
                    DateTime dateObj = DateTime.ParseExact(dateStr, "yyyy-MM-dd", CultureInfo.InvariantCulture);

                    // 設定開始時間和結束時間
                    DateTime start_time = dateObj.Date.AddMinutes(15);
                    DateTime end_time = dateObj.Date.AddDays(1);

                    // 創建時間戳記列表
                    List<string> time_stamps = new List<string>();
                    DateTime current_time = start_time;
                    while (current_time <= end_time)
                    {
                        time_stamps.Add(current_time.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                        current_time = current_time.AddMinutes(15);
                    }

                    foreach (string timeStamp in time_stamps)
                    {
                        Console.WriteLine($"Time Stamp: {timeStamp}");
                    }

                    Console.WriteLine($"FactoryID: {month}, Username: {date}, Password: {password}");

                    Thread.Sleep(5000);

                    if (!ClickElement(driver, "//*[@id=' carrusel']/div/div[1]/div/div[5]/div/a"))
                        continue;

                    Thread.Sleep(5000);
                    // 定義你的下拉選單的 XPath 位置
                    string month_dropdown = "//*[@id='tab5']/div[1]/ul/li[1]/select[2]";
                    string date_dropdown = "//*[@id='tab5']/div[1]/ul/li[1]/select[3]";

                    // 選擇月份和日期
                    SelectDropdownOption(driver, month_dropdown, month.ToString());
                    Thread.Sleep(5000);
                    SelectDropdownOption(driver, date_dropdown, date.ToString());
                    Thread.Sleep(5000);

                    if (!ClickElement(driver, "//*[@id='tab5']/div[1]/ul/li[2]/input"))
                        continue;

                    Thread.Sleep(10000);

                    string pageSource = driver.PageSource;

                    // 提取数据
                    Dictionary<string, List<string>> _data = new Dictionary<string, List<string>>();

                    // 離峰時段
                    var match1 = Regex.Match(pageSource, "highchart_x11 =(.*?);");
                    var submatches1 = Regex.Matches(match1.Groups[1].Value, @"\d+|null");
                    var z = submatches1.OfType<Match>().Select(m => m.Value).ToList();

                    int count = z.Count(s => s == "0000");
                    for (int ioi = 0; ioi < count; ioi++)
                    {
                        z.Remove("7");
                        z.Remove("0000");
                    }

                    _data["off_peak"] = z;



                    foreach (var pair in _data)
                    {
                        Console.WriteLine($"Key: {pair.Key}, Value: [{string.Join(", ", pair.Value)}]");
                    }

                    // 半尖峰時段和週六半尖峰時段
                    var match2 = Regex.Match(pageSource, "highchart_x12 =(.*?);");
                    var submatches2 = Regex.Matches(match2.Groups[1].Value, @"\d+|null");
                    var x = submatches2.OfType<Match>().Select(m => m.Value).ToList();

                    int countX = x.Count(s => s == "0000");
                    for (int iiii = 0; iiii < countX; iiii++)
                    {
                        x.Remove("7");
                        x.Remove("0000");
                    }

                    _data["half_rush_Saturday"] = x;

                    // rush名稱
                    var match3 = Regex.Match(pageSource, "highchart_titleName2 =(.*?);");
                    string rr = match3.Groups[1].Value.Replace("'", "").Replace(" ", "");
                    _data["half_rush_sp"] = new List<string>() { rr };

                    // 尖峰時段
                    var match4 = Regex.Match(pageSource, "highchart_x13 =(.*?);");
                    var submatches4 = Regex.Matches(match4.Groups[1].Value, @"\d+|null");
                    var c = submatches4.OfType<Match>().Select(m => m.Value).ToList();

                    int countC = c.Count(s => s == "0000");
                    for (int ijk = 0; ijk < countC; ijk++)
                    {
                        c.Remove("7");
                        c.Remove("0000");
                    }

                    _data["rush_hour"] = c;

                    // 從字典中獲取值
                    List<string> off_peak = _data["off_peak"];
                    List<string> half_rush_Saturday = _data["half_rush_Saturday"];
                    List<string> half_rush_sp = _data["half_rush_sp"];
                    List<string> rush_hour = _data["rush_hour"];


                    Dictionary<string, Dictionary<string, string>> values = new Dictionary<string, Dictionary<string, string>>();

                    // Traverse the timestamp list and update the values dictionary
                    for (int k = 0; k < time_stamps.Count; k++)
                    {
                        string ts = time_stamps[k];

                        if (off_peak.Count > k && off_peak[k] != "null")
                        {
                            values[ts] = new Dictionary<string, string> { { "name", "離峰時段" }, { "value", off_peak[k] } };
                        }
                        else if (half_rush_Saturday.Count > k && half_rush_Saturday[k] != "null")
                        {
                            string half_rush_name = half_rush_sp.Count > 0 ? half_rush_sp[0] : "";
                            if (half_rush_name == "週六半尖峰")
                            {
                                half_rush_name = "週六半尖峰時段";
                            }

                            values[ts] = new Dictionary<string, string> { { "name", half_rush_name }, { "value", half_rush_Saturday[k] } };
                        }
                        else if (rush_hour.Count > k && rush_hour[k] != "null")
                        {
                            values[ts] = new Dictionary<string, string> { { "name", "尖峰時段" }, { "value", rush_hour[k] } };
                        }
                    }


                    foreach (var kvp in values)
                    {
                        string timestamp = kvp.Key;
                        string name = kvp.Value["name"];
                        string value = kvp.Value["value"];

                        Console.WriteLine($"Timestamp: {timestamp}");
                        Console.WriteLine($"Name: {name}");
                        Console.WriteLine($"Value: {value}");
                    }

                    foreach (var kvp in values)
                    {
                        string timestamp = kvp.Key;
                        string name = kvp.Value["name"];
                        string value = kvp.Value["value"];
                        Console.WriteLine($"Value: {value}");

                        if (name == "週六半尖峰")
                        {
                            name = "週六半尖峰時段";
                        }

                        string sql = $@"
                            INSERT INTO Test_TPC
                            VALUES ('{factoryid}', LEFT('{timestamp}', 23), '{name}', {value})
                        ";

                        // 建立連線物件
                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            // 開啟連線
                            connection.Open();

                            // 建立命令物件
                            using (SqlCommand command = new SqlCommand(sql, connection))
                            {
                                try
                                {
                                    // 執行 SQL 指令碼
                                    command.ExecuteNonQuery();
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"An error occurred: {ex.Message}");
                                }
                            }
                        }
                    }


                }

                Thread.Sleep(2000);
            }

            Thread.Sleep(500000000);
            // 关闭 ChromeDriver
            driver.Quit();

            // 等待用户输入，防止控制台立即关闭
            Console.ReadLine();
        }

        static void SelectDropdownOption(IWebDriver driver, string xpath, string value)
        {
            // 找到下拉選單元素
            IWebElement dropdown = driver.FindElement(By.XPath(xpath));

            // 建立一個 SelectElement
            var selectElement = new SelectElement(dropdown);

            // 選擇下拉選單的選項
            selectElement.SelectByValue(value);
        }

        static bool ClickElement(IWebDriver driver, string xpath)
        {
            try
            {
                Console.WriteLine(xpath);
                IWebElement element = FindElement(driver, By.XPath(xpath), 15);
                element.Click();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to click element: {xpath}");
                Console.WriteLine(ex.ToString());
                return false;
            }
        }

        static IWebElement FindElement(IWebDriver driver, By by, int timeoutInSeconds)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutInSeconds));
                try
                {
                    return wait.Until(drv =>
                    {
                        var element = drv.FindElement(by);
                        return (element.Displayed && element.Enabled) ? element : null;
                    });
                }
                catch (WebDriverTimeoutException)
                {
                    return null;
                }
            }

            return driver.FindElement(by);
        }

        static Func<IWebDriver, IAlert> AlertIsPresent()
        {
            return (driver) =>
            {
                try
                {
                    return driver.SwitchTo().Alert();
                }
                catch (NoAlertPresentException)
                {
                    return null;
                }
            };
        }
    }

    public class DatabaseReader
    {
        private string _connectionString;

        public DatabaseReader(string connectionString)
        {
            _connectionString = connectionString;
        }

        public List<Tuple<string, string, string>> GetFactoryData()
        {
            var result = new List<Tuple<string, string, string>>();

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                string sql = "SELECT FactoryID, FactoryName, Tpc_act, Tpc_pwd FROM Factory WHERE Tpc_act IS NOT NULL ORDER BY aOrder";
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string factoryid = reader.GetString(0); // 對應到 Tpc_act
                            string username = reader.GetString(2); // 對應到 Tpc_act
                            string password = reader.GetString(3); // 對應到 Tpc_pwd
                            result.Add(Tuple.Create(factoryid, username, password));
                        }
                    }
                }
            }

            return result;
        }
    }
}
