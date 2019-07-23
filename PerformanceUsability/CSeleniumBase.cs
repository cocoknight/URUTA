/*********************************************************************************************************-- 
    
    Copyright (c) 2018, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

    2019-12-10   : Make a Selenium Base
    
    2019-02-09 : Dictionary초기화 코드. 추가 각 Task가 다시 수행되었을때 이전 Dictionary를 처리해 주지 않으면 Duplicate Key에러 발생함
    2019-02-16 : TaskStartReport,TaskFinishReport 메서드 Data와 View형태로 분리
    TaskStartReport --->TaskUpdateData, TaskUpdateView. 
    -add resetDictionaryKey member funtion
    2019-03-08 : Add browser options before starting browser
    2019-03-18 : Handle Explicit Wait Code
    2019-03-19 : WebDriver Start Option에 시간설정 추가
    2019-03-20 : chrome no-sandbox option추가 함.
    2019-04-03 : Fix Math.Round calulation for discharge_wh
    2019-04-03 : Type change double to integer for BatteryLife Value   
    2019-04-03 : Fix Math.Round calulation for running time
    2019-04-04 : After Task Finish, Update Running Time
    2019-06-06 : Add Exceptoin Handling for initSelenium with Retry-Handler.

    //Selenium and Chrome Driver Exception Handling.
    https://github.com/SeleniumHQ/selenium/issues/6317
    https://groups.google.com/forum/#!topic/selenium-users/jM6yyeNKYrI 
    https://stackoverflow.com/questions/48232737/selenium-driver-quit-not-working-when-passed-by-reference


    //packaeg 설치폴더
    //C:\Users\ymkim\AppData\Local\Apps\2.0\EG7WGWHP.Z47\DB6EZ705.3ZY\perf..tion_ff61be73852145bc_0001.0000_6c26f63f4e2bc74e
--***********************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


//Seleniu Test Part
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Edge;
using System.IO;
using System.Windows.Forms;

namespace PerformanceUsability
{
    class CSeleniumBase
    {
        //Data Member
        //접근 권한을 명시적으로 지정하지 않으면 default로 private 권한이 지정된다.
        //private로 지정되면 상속클래스에도 직접 data-member를 접근할 수 없다.
        //protected이상 권한이면 접근 가능하다.

        //public WebType _webType { get; set; }
        protected WebType _webType { get; set; }
        protected IWebDriver _driver;
        //protected Form1 _uiManager;
        public Form1 _uiManager;
        //TOAN : 12/25/2018
        //JavascriptExecutor _js;


        //Below is Singletone Class for KeyList
        protected KeyList _keyList;
        protected CUtility _myUtility;
        public Dictionary<string, string> _columnInfoDic;

        //Default Constructor
        //TOAN : 04/03/2019. 타입 변경
        //protected double _startBattery;
        //double _remaining_battery;
        //double _discharge;
        protected int _startBattery;
        int _remaining_battery;
        int _discharge;

        double _discharge_wh;
        double _powerConsumption;

        protected DateTime _taskStartTime;
        protected DateTime _taskEndTime;


        protected string s_task_number;
        protected string s_testcase;
        protected string s_remaining_battery;
        protected string s_status;
        protected string s_discharge;
        protected string s_discharge_wh;
        protected string s_power_consumption_wh;
        protected string s_startTime;
        protected string s_endTime;
        //TOAN : 04/04/2019.
        protected string s_runningTime;
        public int retryCounter;

        //TOAN : 07/03/2019. Test Terminate Condition
       // public bool _testTerminate;

        public bool _testTerminate { get; set; }

        public CSeleniumBase()
        {
            //System.Diagnostics.Debug.WriteLine("{0}:Default Constructor");
            System.Diagnostics.Debug.WriteLine("CSeleniumBase");

            _keyList = KeyList.Instance;
            _myUtility = CUtility.Instance;

            string currTC = _keyList.k_testcase;
            string currNo = _keyList.k_testcase_no;
            _columnInfoDic = new Dictionary<string, string>();

            _testTerminate = false;
        }
        //Constructor
        public CSeleniumBase(WebType type)
        {
            System.Diagnostics.Debug.WriteLine("CSeleniumBase with WebType");
            _webType = type;
            //this.initSelenium(type);

            _keyList = KeyList.Instance;
            _myUtility = CUtility.Instance;

            string currTC = _keyList.k_testcase;
            string currNo = _keyList.k_testcase_no;
            _columnInfoDic = new Dictionary<string, string>();

            //TOAN : 06/06/2019. set selenium retrycounter
            retryCounter = 3;
            _testTerminate = false;
        }

        public void setTimeWait(int sec)
        {
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(sec);
        }

        public void initSelenium(WebType mode)
        {
              try
                {
                //case1
                string driverPath = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                //System.Diagnostics.Debug.WriteLine(string.Format("execution path: {0}", driverPath));
                //IWebDriver driver = new ChromeDriver(driverPath);

                //case2
                //string cPath = System.Reflection.Assembly.GetExecutingAssembly().Location; //executable path
                //string appPath = Path.GetDirectoryName(Application.ExecutablePath); //executable directory path
                                                                                    //System.Diagnostics.Debug.WriteLine(string.Format("execution path: {0}", appPath));

                //_driver = new OpenQA.Selenium.Chrome.ChromeDriver();


                //case3
                string appPath =System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                System.Diagnostics.Debug.WriteLine(string.Format("execution path: {0}", appPath));

                //TOAN :06/25/2019. add option
                var options = new ChromeOptions();
                options.AddAdditionalCapability("useAutomationExtension", false);
                options.AddArguments("--ignore-certificate-errors");
                options.AddArguments("--ignore-ssl-errors");

                //TOAN : 06/25/2019. increate timespan 1->2 minute
                //If we don't put value. Default value is 1 minute
                _driver = new OpenQA.Selenium.Chrome.ChromeDriver(appPath,options, TimeSpan.FromMinutes(2));
                //_driver = new OpenQA.Selenium.Chrome.ChromeDriver()
                //_driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100/*50*/);

                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                   
                }
                finally
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Finally Block Running"));
                }
           

            //   switch (mode)
            //   {
            //    case WebType.WEB_Chrome:
            //        {

            //            //TOAN :  03/08/2019.
            //            //var options = new ChromeOptions();
            //            //options.AddAdditionalCapability("useAutomationExtension", false);
            //            //options.AddArguments("--ignore-certificate-errors");
            //            //options.AddArguments("--ignore-ssl-errors");

            //            ////TOAN : 03/16/2019. Chrome ignore save password pop-up
            //            //options.AddUserProfilePreference("credentials_enable_service", false);
            //            //options.AddUserProfilePreference("profile.password_manager_enabled", false);

            //            //TOAN : 03/20/2019. Web Driver timeout after 60 seconds
            //            //options.AddArgument("no-sandbox");


            //            //TOAN : 03/19/2019. driver생성 시 timespan추가해서 진행
            //            //_driver = new ChromeDriver(options);
            //            _driver = new ChromeDriver();


            //            //TimeSpan이 포함된 생성자를 호출해야 한다.
            //            //ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            //            //service.SuppressInitialDiagnosticInformation = true;
            //            //_driver = new ChromeDriver(service, options, TimeSpan.FromMinutes(3));

            //            //TOAN : 03/18/2019. applied to wait all sesseion with implicit waite. increase wait(300->500)
            //            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);
            //            //_driver = new ChromeDriver();

            //            break;
            //        }
            //    case WebType.WEB_FireFox:
            //        {
            //            _driver = new FirefoxDriver();
            //            break;
            //        }
            //    case WebType.WEB_IE:
            //        {
                        
            //            _driver = new InternetExplorerDriver();
            //            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(500);
            //            break;
            //        }
            //    case WebType.WEB_EDGE:
            //        {
            //            _driver = new EdgeDriver();
            //            break;
            //        }
            //    default:
            //        break;
            //}
        }

        public void connectUI(Form1 conn)
        {
            _uiManager = conn;
            System.Diagnostics.Debug.WriteLine("connectUI(CSeleniumBase)");
            conn.HeyConnect();
        }
        //C#의 경우 별도로 소멸자(Destructor)를 구현하지 않는다.
        //JAVA와 유사하게 CLR(Common Language Runtime)의 가비지 컬렉터가 객체가 소멸되는 시점을
        //판단해 소멸자를 호출 한다. CLR의 가비지 컬렉터는 명시적으로 프로그래머가 소멸자를 구현하는것 보다
        //훨씬 똑똑하게 객체의 소멸을 처리 한다.
        /*
        ~CSeleniumBase()
        {

        }
        */
        public void updateTaskResult(TaskStatus status)
        {
            Dictionary<string, string> itemInfo = _columnInfoDic;

            //protected string s_task_number;
            //protected string s_testcase;
            //protected string s_remaining_battery;
            //protected string s_status;
            //protected string s_discharge;
            //protected string s_discharge_wh;
            //protected string s_power_consumption_wh;
            //protected string s_startTime;
            //protected string s_endTime;


            switch (status)
            {
                case TaskStatus.TASK_RUNNING:
                    {
                        System.Diagnostics.Debug.WriteLine("TASK_RUNNING");
                        this.resetDictionaryKey(itemInfo, TaskStatus.TASK_RUNNING);
                        ////기존 키가 존재했다면 해당 키를 지우고 시작한다.
                        //if (itemInfo.ContainsKey(_keyList.k_testcase_no))
                        //{
                        //    itemInfo.Remove(_keyList.k_testcase_no);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_testcase))
                        //{
                        //    itemInfo.Remove(_keyList.k_testcase);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_status))
                        //{
                        //    itemInfo.Remove(_keyList.k_status);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_start_time))
                        //{
                        //    itemInfo.Remove(_keyList.k_start_time);
                        //}

                        //시작 전 컬럼별로 상태값을 가지고 온다.
                        //_remaining_battery = _myUtility.getBatteryLife();
                        _taskStartTime = System.DateTime.Now;
                        //TOAN : 04/03/2019. Double->Integer값으로 진행(double값 계산시 0.999999999와 같은 상황 발생)
                        //_startBattery = _myUtility.getBatteryLife();
                        _startBattery = _myUtility.getBatteryLifeV1();

                        //시작 전 상태를 string변수에 기록한다.
                        s_task_number = this.testcase_no;
                        // s_remaining_battery = _remaining_battery.ToString() + "%";
                        s_testcase = this.testcase_name;
                        s_status = "Running";
                        //s_startTime = _taskStartTime.ToString();
                        s_startTime = string.Format("{0:hh:mm tt}", _taskStartTime);


                        //compose return value
                        itemInfo.Add(_keyList.k_testcase_no, s_task_number);
                        itemInfo.Add(_keyList.k_testcase, s_testcase);
                        itemInfo.Add(_keyList.k_status, s_status);
                        //itemInfo.Add(_keyList.k_remaining_battery, s_remaining_battery);
                        itemInfo.Add(_keyList.k_start_time, s_startTime);
                        break;
                    }
                case TaskStatus.TASK_FINISH:
                    {
                        this.resetDictionaryKey(itemInfo, TaskStatus.TASK_FINISH);
                        //기존키가 존재했다면 해당 키를 지우고 시작한다.
                        //if (itemInfo.ContainsKey(_keyList.k_remaining_battery))
                        //{
                        //    itemInfo.Remove(_keyList.k_remaining_battery);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_discharge))
                        //{
                        //    itemInfo.Remove(_keyList.k_discharge);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_discharge_wh))
                        //{
                        //    itemInfo.Remove(_keyList.k_discharge_wh);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_power_consumption_wh))
                        //{
                        //    itemInfo.Remove(_keyList.k_power_consumption_wh);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_end_time))
                        //{
                        //    itemInfo.Remove(_keyList.k_end_time);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_status))
                        //{
                        //    itemInfo.Remove(_keyList.k_status);
                        //}


                        System.Diagnostics.Debug.WriteLine("TASK_FINISH");


                        //종료 후 상태를 ListView에 기록한다.
                        _taskEndTime = System.DateTime.Now;
                        //TOAN : 04/03/2019. Double->Integer값으로 진행(double값 계산시 0.999999999와 같은 상황 발생)
                        //_remaining_battery = _myUtility.getBatteryLife();
                        _remaining_battery = _myUtility.getBatteryLifeV1();

                        //TOAN : 02/10/2019. Below is new code
                        _discharge = _startBattery - _remaining_battery;
                        //_discharge = Math.Round(_discharge / 100, 2);
                        //TOAN : 04/03/2019. 수식 계산순서 변경
                        //_discharge_wh = _myUtility.getBatteryWH() * Math.Round(_discharge / 100, 2);
                        //_discharge_wh = Math.Round(_discharge_wh, 2);
                        
                        //TOAN : 04/03/2019. _discharge는 int형이므로 소수 계산을 위해 double형으로 변환
                        _discharge_wh = Math.Round(_myUtility.getBatteryWH() * (Convert.ToDouble(_discharge) / 100), 1, MidpointRounding.AwayFromZero);


                        //TOAN : 02/10/2019. Below is original code
                        //_discharge = _startBattery - _remaining_battery;
                        //_discharge_wh = _myUtility.getBatteryWH() * (_discharge / 100);
                        //_discharge_wh = Math.Round(_discharge_wh, 2);

                        //calculate power consumption
                        double calTotalMninutes = _taskEndTime.Subtract(_taskStartTime).TotalMinutes;
                        System.Diagnostics.Debug.WriteLine("Total Minutes:{0}", calTotalMninutes);

                        double calToHour = calTotalMninutes / 60;
                        System.Diagnostics.Debug.WriteLine("Total Hour:{0}", calToHour);

                        //TOAN : 04/04/2019. 소수점 3자리에서 반올림
                        //double convertHour = Math.Round(calToHour, 2);   //소수점 3째자리에서 반올림, 2째자리까지만 유효하게 한다.
                        double convertHour = Math.Round(calToHour, 2, MidpointRounding.AwayFromZero);
                        System.Diagnostics.Debug.WriteLine("Total Convert Hour:{0}", convertHour);

                        //TOAN : 04/03/2019. power consumption코드 변경.
                        //_powerConsumption = Math.Round(_discharge_wh / convertHour, 2);
                        _powerConsumption = Math.Round(_discharge_wh / convertHour, 1, MidpointRounding.AwayFromZero);
                        //compose return value
                        s_remaining_battery = _remaining_battery.ToString() + "%";
                        //s_endTime = _taskEndTime.ToString();
                        s_status = "Finish";
                        s_endTime = string.Format("{0:hh:mm tt}", _taskEndTime);

                        //TOAN : 04/04/2019
                        s_runningTime = convertHour.ToString() + "hr";

                        //TOAN : 02/18/2019. NaN처리
                        if (double.IsNaN(_discharge))
                        {
                            s_discharge = 0.ToString() + "%";
                        }
                        else
                        {
                            s_discharge = _discharge.ToString() + "%";
                        }


                        if (double.IsNaN(_discharge_wh))
                        {
                            s_discharge_wh = 0.ToString() + "Wh";
                        }
                        else
                        {
                            s_discharge_wh = _discharge_wh.ToString() + "Wh";
                        }

                        if(double.IsNaN(_powerConsumption))
                        {
                            s_power_consumption_wh = 0.ToString() + "W";
                        }
                        else
                        {
                            s_power_consumption_wh = _powerConsumption.ToString() + "W";
                        }
                        //s_discharge = _discharge.ToString() + "%";
                        //s_discharge_wh = _discharge_wh.ToString() + "Wh";
                        //s_power_consumption_wh = _powerConsumption.ToString() + "W";
                        
                        itemInfo.Add(_keyList.k_remaining_battery, s_remaining_battery);
                        itemInfo.Add(_keyList.k_discharge, s_discharge);
                        itemInfo.Add(_keyList.k_discharge_wh, s_discharge_wh);
                        itemInfo.Add(_keyList.k_power_consumption_wh, s_power_consumption_wh);
                        itemInfo.Add(_keyList.k_end_time, s_endTime);
                        itemInfo.Add(_keyList.k_status, s_status);
                        //TOAN  04/04/2019. add running time
                        itemInfo.Add(_keyList.k_running_time, s_runningTime);

                        break;
                    }

                default:
                    break;
            }
            
        }

        public Dictionary<string, string> composeTaskResult(TaskStatus status)
        {
            Dictionary<string, string> itemInfo = _columnInfoDic;

           //protected string s_task_number;
           //protected string s_testcase;
           //protected string s_remaining_battery;
           //protected string s_status;
           //protected string s_discharge;
           //protected string s_discharge_wh;
           //protected string s_power_consumption_wh;
           //protected string s_startTime;
           //protected string s_endTime;

            
            switch (status)
            {
                case TaskStatus.TASK_RUNNING:
                    {
                        System.Diagnostics.Debug.WriteLine("TASK_RUNNING");
                        this.resetDictionaryKey(itemInfo, TaskStatus.TASK_RUNNING);
                        ////기존 키가 존재했다면 해당 키를 지우고 시작한다.
                        //if (itemInfo.ContainsKey(_keyList.k_testcase_no))
                        //{
                        //    itemInfo.Remove(_keyList.k_testcase_no);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_testcase))
                        //{
                        //    itemInfo.Remove(_keyList.k_testcase);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_status))
                        //{
                        //    itemInfo.Remove(_keyList.k_status);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_start_time))
                        //{
                        //    itemInfo.Remove(_keyList.k_start_time);
                        //}

                        //시작 전 컬럼별로 상태값을 가지고 온다.
                        //_remaining_battery = _myUtility.getBatteryLife();
                        _taskStartTime = System.DateTime.Now;
                        _startBattery = _myUtility.getBatteryLifeV1();

                        //시작 전 상태를 string변수에 기록한다.
                        s_task_number = this.testcase_no;
                       // s_remaining_battery = _remaining_battery.ToString() + "%";
                        s_testcase = this.testcase_name;
                        s_status = "Running";
                        //s_startTime = _taskStartTime.ToString();
                        s_startTime = string.Format("{0:hh:mm tt}", _taskStartTime);

                        
                       //compose return value
                        itemInfo.Add(_keyList.k_testcase_no, s_task_number);
                        itemInfo.Add(_keyList.k_testcase, s_testcase);
                        itemInfo.Add(_keyList.k_status, s_status);
                        //itemInfo.Add(_keyList.k_remaining_battery, s_remaining_battery);
                        itemInfo.Add(_keyList.k_start_time, s_startTime);
                         break;
                    }
                case TaskStatus.TASK_FINISH:
                    {
                        this.resetDictionaryKey(itemInfo, TaskStatus.TASK_FINISH);
                        //기존키가 존재했다면 해당 키를 지우고 시작한다.
                        //if (itemInfo.ContainsKey(_keyList.k_remaining_battery))
                        //{
                        //    itemInfo.Remove(_keyList.k_remaining_battery);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_discharge))
                        //{
                        //    itemInfo.Remove(_keyList.k_discharge);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_discharge_wh))
                        //{
                        //    itemInfo.Remove(_keyList.k_discharge_wh);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_power_consumption_wh))
                        //{
                        //    itemInfo.Remove(_keyList.k_power_consumption_wh);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_end_time))
                        //{
                        //    itemInfo.Remove(_keyList.k_end_time);
                        //}

                        //if (itemInfo.ContainsKey(_keyList.k_status))
                        //{
                        //    itemInfo.Remove(_keyList.k_status);
                        //}


                        System.Diagnostics.Debug.WriteLine("TASK_FINISH");
                       

                        //종료 후 상태를 ListView에 기록한다.
                        _taskEndTime = System.DateTime.Now;
                        _remaining_battery = _myUtility.getBatteryLifeV1();

                        //TOAN : 02/10/2019. Below is new code
                        //TOAN : 04/03/2019. code optimize
                        _discharge = _startBattery - _remaining_battery;
                        _discharge_wh = Math.Round(_myUtility.getBatteryWH() * (Convert.ToDouble(_discharge) / 100), 1, MidpointRounding.AwayFromZero);
                        //_discharge_wh = _myUtility.getBatteryWH() * Math.Round(_discharge / 100, 2);
                        //_discharge_wh = Math.Round(_discharge_wh, 2);



                        //TOAN : 02/10/2019. Below is original code
                        //_discharge = _startBattery - _remaining_battery;
                        //_discharge_wh = _myUtility.getBatteryWH() * (_discharge / 100);
                        //_discharge_wh = Math.Round(_discharge_wh, 2);

                        //calculate power consumption
                        double calTotalMninutes = _taskEndTime.Subtract(_taskStartTime).TotalMinutes;
                        System.Diagnostics.Debug.WriteLine("Total Minutes:{0}",calTotalMninutes);

                        double calToHour = calTotalMninutes / 60;
                        System.Diagnostics.Debug.WriteLine("Total Hour:{0}", calToHour);

                        //TOAN : 04/04/2019. 소수점 3자리에서 반올림
                        //double convertHour = Math.Round(calToHour, 2);   //소수점 3째자리에서 반올림, 2째자리까지만 유효하게 한다.
                        double convertHour = Math.Round(calToHour, 2, MidpointRounding.AwayFromZero);
                        System.Diagnostics.Debug.WriteLine("Total Convert Hour:{0}", convertHour);

                        //TOAN : 04/03/2019. power consumption코드 변경.
                        //_powerConsumption = Math.Round(_discharge_wh / convertHour,2);
                        _powerConsumption = Math.Round(_discharge_wh / convertHour, 1, MidpointRounding.AwayFromZero);
                        //compose return value

                        s_remaining_battery = _remaining_battery.ToString() + "%";
                        //s_endTime = _taskEndTime.ToString();
                        s_status = "Finish";
                        s_endTime = string.Format("{0:hh:mm tt}", _taskEndTime);

                        //TOAN : 04/04/2019
                        s_runningTime = convertHour.ToString() + "hr";

                        s_discharge = _discharge.ToString() + "%";
                        s_discharge_wh = _discharge_wh.ToString() + "Wh";
                        s_power_consumption_wh = _powerConsumption.ToString() + "W";


                        itemInfo.Add(_keyList.k_remaining_battery, s_remaining_battery);
                        itemInfo.Add(_keyList.k_discharge, s_discharge);
                        itemInfo.Add(_keyList.k_discharge_wh, s_discharge_wh);
                        itemInfo.Add(_keyList.k_power_consumption_wh, s_power_consumption_wh);
                        itemInfo.Add(_keyList.k_end_time, s_endTime);
                        itemInfo.Add(_keyList.k_status, s_status);
                        //TOAN  04/04/2019. add running time
                        itemInfo.Add(_keyList.k_running_time, s_runningTime);
                        break;
                    }

                default:
                    break;
            }
            return itemInfo;
        }
        //public Dictionary<string, string> composeTaskResult(TaskStatus status)
        //{
        //    Dictionary<string, string> itemInfo = _columnInfoDic;
        //    //_keyList
        //    //접근 제한자를 명시하지 않으면 default private이다
        //    double remaining_battery;
        //    double discharge;
        //    double discharge_wh;
        //    double powerConsumption;

        //    string s_task_number;
        //    string s_testcase;
        //    string s_remaining_battery;
        //    string s_status;
        //    string s_discharge;
        //    string s_discharge_wh;
        //    string s_power_consumption_wh;
        //    string s_startTime;
        //    string s_endTime;



        //    switch (status)
        //    {

        //        default:
        //            break;

        //    }
        //    //s_task_number = Int32.Parse(batterylife); ;

        //    //종료시간
        //    s_endTime = _myUtility.getCurrentTime();
        //    //현재 배터리
        //    remaining_battery = _myUtility.getBatteryLife();
        //    //Task시작의 배터리 상태에서 Task종료된 배터리 값 차이구하면 이 값이 discharge이다. 
        //    discharge = _startBattery - remaining_battery;
        //    discharge_wh = _myUtility.getBatteryWH() * (discharge / 100);
        //    discharge_wh = Math.Round(discharge_wh, 2);
        //    //powerConsumption = discharge_wh/

        //    s_remaining_battery = remaining_battery.ToString() + "%";
        //    s_discharge = discharge.ToString() + "%";
        //    s_discharge_wh = discharge_wh.ToString() + "Wh";



        //    //Task Number는 Class instance를 생성할 때, 생성자로 받아오던지, 아니면 별도 멤버펑션으로 추가 한다.
        //    //s_task_number = "1";  
        //    //s_status = "완료";
        //    //s_remaining_battery = "95%";
        //    //s_discharge = "5%";
        //    //s_discharge_wh = "3.8Wh";
        //    //s_power_consumption_wh = "7.5Wh";
        //    //s_startTime = "4:11 PM";
        //    //s_endTime = "4.14 PM";

        //    //System.Diagnostics.Debug.WriteLine(s_task_number);
        //    //System.Diagnostics.Debug.WriteLine(s_status);


        //    //itemInfo.Add(_keyList.k_testcase_no, s_task_number);
        //    //itemInfo.Add(_keyList.k_status, s_status);
        //    itemInfo.Add(_keyList.k_remaining_battery, s_remaining_battery);
        //    itemInfo.Add(_keyList.k_discharge, s_discharge);
        //    itemInfo.Add(_keyList.k_discharge_wh, s_discharge_wh);
        //    //itemInfo.Add(_keyList.k_power_consumption_wh, s_power_consumption_wh);
        //    //itemInfo.Add(_keyList.k_start_time, s_startTime);
        //    itemInfo.Add(_keyList.k_end_time, s_endTime);

        //    return itemInfo;
        //}

        public string testcase_name
        {
            get;
            set;
        }

        public string testcase_no
        {
            get;
            set;
        }

        public void checkTaskStartCondition()
        {
          
        }

        public void checkTaskEndCondition()
        {

        }

        public void TaskFinishReport()
        {
            Dictionary<string, string> taskResult = this.composeTaskResult(TaskStatus.TASK_FINISH);
            _uiManager.HandleTaskReport(taskResult, TaskStatus.TASK_FINISH);
        }

        public void TaskUpdateData(TaskStatus status)
        {
            this.updateTaskResult(status);
        }
            
        public void TaskUpdateView(TaskStatus status)
        {
            _uiManager.HandleTaskReport(_columnInfoDic, status);
        }

        public void TaskStartReport()
        {
            Dictionary<string, string> taskResult = this.composeTaskResult(TaskStatus.TASK_RUNNING);
            _uiManager.HandleTaskReport(taskResult, TaskStatus.TASK_RUNNING);
        }

        public void TaskRunningRecord(TaskRunningList runningTask)
        {
            //uiManager
            _uiManager.recordRunningList(runningTask);
        }


        public void resetDictionaryKey(Dictionary<string, string> cList,TaskStatus status)
        {

            switch (status)
            {
                 
                case TaskStatus.TASK_RUNNING:
                    {

                        //기존 키가 존재했다면 해당 키를 지우고 시작한다.
                        cList.Clear();
                        //if (cList.ContainsKey(_keyList.k_testcase_no))
                        //{
                        //    cList.Remove(_keyList.k_testcase_no);
                        //}

                        //if (cList.ContainsKey(_keyList.k_testcase))
                        //{
                        //    cList.Remove(_keyList.k_testcase);
                        //}

                        //if (cList.ContainsKey(_keyList.k_status))
                        //{
                        //    cList.Remove(_keyList.k_status);
                        //}

                        //if (cList.ContainsKey(_keyList.k_start_time))
                        //{
                        //    cList.Remove(_keyList.k_start_time);
                        //}

                        //if (cList.ContainsKey(_keyList.k_remaining_battery))
                        //{
                        //    cList.Remove(_keyList.k_remaining_battery);
                        //}

                        //if (cList.ContainsKey(_keyList.k_discharge))
                        //{
                        //    cList.Remove(_keyList.k_discharge);
                        //}

                        //if (cList.ContainsKey(_keyList.k_discharge_wh))
                        //{
                        //    cList.Remove(_keyList.k_discharge_wh);
                        //}

                        //if (cList.ContainsKey(_keyList.k_power_consumption_wh))
                        //{
                        //    cList.Remove(_keyList.k_power_consumption_wh);
                        //}

                        //if (cList.ContainsKey(_keyList.k_end_time))
                        //{
                        //    cList.Remove(_keyList.k_end_time);
                        //}

                        //if (cList.ContainsKey(_keyList.k_status))
                        //{
                        //    cList.Remove(_keyList.k_status);
                        //}

                        break;
                    }
                case TaskStatus.TASK_FINISH:
                    {

                        if (cList.ContainsKey(_keyList.k_remaining_battery))
                        {
                            cList.Remove(_keyList.k_remaining_battery);
                        }

                        if (cList.ContainsKey(_keyList.k_discharge))
                        {
                            cList.Remove(_keyList.k_discharge);
                        }

                        if (cList.ContainsKey(_keyList.k_discharge_wh))
                        {
                            cList.Remove(_keyList.k_discharge_wh);
                        }

                        if (cList.ContainsKey(_keyList.k_power_consumption_wh))
                        {
                            cList.Remove(_keyList.k_power_consumption_wh);
                        }

                        if (cList.ContainsKey(_keyList.k_end_time))
                        {
                            cList.Remove(_keyList.k_end_time);
                        }

                        if (cList.ContainsKey(_keyList.k_status))
                        {
                            cList.Remove(_keyList.k_status);
                        }

                        //TOAN : 04/04/2019. add k_running_time
                        if (cList.ContainsKey(_keyList.k_running_time))
                        {
                            cList.Remove(_keyList.k_running_time);
                        }

                        break;
                    }
                    default:
                    break;
            }

        }

        public IWebDriver getWebDriver()
        {
            return _driver;
        }

        //TOAN : 07/02/2019. return next node
        public LinkedListNode<TaskRunningList> getNextTask(LinkedListNode<TaskRunningList> currNode)
        {
            LinkedListNode<TaskRunningList> node=null;

            //check currentNode is lastnode or not
            //node=_uiManager.getTaskList().Find(TaskRunningList.TASK_WEBACTOR);
            node = _uiManager.getTaskList().Last;
            if(node.Value==currNode.Value)
            {
                //currNode is last
                node = _uiManager.getTaskList().First;
            }
            else
            {
                //currNode is not last
                node = currNode.Next;
            }
            return node;
        }

    }


   
}
