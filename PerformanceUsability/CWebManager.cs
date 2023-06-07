/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

  2019-06-30 : Add new Web automation Manager class
  2019-08-22 : Exception발생 후, 기존 URL Retry시 driver호출코드가 빠져 있었음. 
               이경우 계속해서 인터넷창이 남아있고, 에러 난것으로 보고 됨. 즉 exception발생이후 retry를 하지 않음.
  2020-09-04 : Naver "최신 인기 영화" DOM변경에 따른 소스 변경

  key word : how to cancel background worker after specified time in c#
  URL : https://stackoverflow.com/questions/1341488/how-to-cancel-background-worker-after-specified-time-in-c-sharp

  key word : Using Timer inside a BackGroundWorker
  URL : https://stackoverflow.com/questions/6704195/using-timer-inside-a-backgroundworker

  key word : Background Worker Cancel
  URL : https://www.wpf-tutorial.com/misc/cancelling-the-backgroundworker/

  key word : web page assert confirm
  URL : https://stackoverflow.com/questions/51282067/how-to-validate-page-title-is-correct-actual-to-expected-selenium-c-sharp
--***********************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



//Selenium Test Part
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Edge;

using System.Threading;
using OpenQA.Selenium.Remote;

//Timer
//using System.Windows.Forms.Timer;
using System.Windows.Forms;
using OpenQA.Selenium.Support.UI;
using System.Globalization;
using System.ComponentModel;
using System.Timers;

namespace PerformanceUsability
{
    class CWebManager : CSeleniumBase
    {
        

        //public WebType _webType { get; set; }
        //public WebType _webType { get; set; }
        public int timer_sec;
        public string _startURL;
        bool exit_flag = false;
        string _saveURL;
       

        System.Windows.Forms.Timer _timer;
        System.Timers.Timer _systemTimer;

        public System.ComponentModel.BackgroundWorker worker;
        public bool _workComplete { get; set; }

        public int _finishTime;

        //TOAN : 07/15/2021. 테스트 지역 정보추가
        public string _currRegion = "";

        public CWebManager(WebType type) : base(type)
        {

            _startURL = @"http://www.naver.com";
            System.Diagnostics.Debug.WriteLine("CWebActor with WebType");
            System.Diagnostics.Debug.WriteLine("Task Terminate:{0}", exit_flag);

            //Declare of Background Worker
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
            _workComplete = false;
        }




        //public void SystemTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        //{
        //    //MessageBox.Show("System Timer Expire!");
        //    _systemTimer.Stop();


        //}

        public void setTestTime(int time)
        {
            _finishTime = time * 60;
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {

            string argument = e.Argument as string;
            this.exit_flag = false;

            try
            {
                switch (argument)
                {
                    case "ACTION_START":
                        {

                            this.setSystemTimer(/*600*/_finishTime);
                            //TOAN : 07/15/2021. WebType에 맞게 변경
                            //this.initSelenium(0);
                            this.initSelenium(_webType);


                            //TOAN : 07/15/2021. Korea/China코드 구분 진행
                            _currRegion = _uiManager.getCurrentRegion();

                            //TOAN : 07/15/2021. logic add. 아래 코드에서 exception처리하지 않으면
                            //그냥task가 종료되어 버린다. Timer에 의해 종료되도록 수정
                            try
                            {
                                //TOAN : 08/04/2021. 아래 timeWait는 무시하자.
                                //this.setTimeWait(100);
                                this.TaskUpdateData(TaskStatus.TASK_RUNNING);
                            }catch(Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                            }
                            //TOAN End

                            this.TaskRunningRecord(TaskRunningList.TASK_WEBACTOR);
                            worker.ReportProgress(1); //View Update

                            //this.setSystemTimer(/*600*/_finishTime);
                            exit_flag = false;


                            //1time query
                            //TOAN : 07/15/2021. logic add. 아래 코드에서 exception처리하지 않으면
                            //그냥task가 종료되어 버린다. Timer에 의해 종료되도록 수정
                            try
                            {
                                //TOAN : 07/15/2021. 기존코드 삭제
                                /*
                                _driver.Url = _startURL;
                                IWebElement q = _driver.FindElement(By.Id("query"));

                                q.SendKeys("최신영화순위");
                                _driver.FindElement(By.Id("search_btn")).Click();
                                */

                                if (_currRegion.Equals("CN"))
                                {
                                    //TOAN : 08/05/2021. URL Change. In case of Baidu, Frequentlly error occurs internally on Automation Testing.
                                    //Perhaps, This is internal baidu logic. So, I'll change url to www.so.com(360)
                                    //_startURL = @"http://www.baidu.com";
                                    _startURL = @"http://www.so.com";

                                    //TOAN : 01/19/2022. chromedriver.exe창을 최소화 
                                    this.minimize_edge_driver();
                                }
                                else
                                {
                                    //TOAN : 01/19/2022. chromedriver.exe창을 최소화 
                                    this.minimize_chrome_driver();
                                    _startURL = @"http://www.naver.com";
                                }

                             
                                //TOAN : 01/07/2022. browser screen 최대화
                                //TOAN : 01/17/2022. browser screen 표준 사이즈
                                //_driver.Manage().Window.Maximize();
                                _driver.Url = _startURL;
                                

                                if (_currRegion.Equals("CN"))
                                {
                                    //TOAN : 08/04/2021. for -testing
                                    //Thread.Sleep(3000);
                                    Thread.Sleep(1000);
                                    //Thread.Sleep(5000);

                                    //TOAN : 08/04/2021. use xpath   
                                    //IWebElement q = _driver.FindElement(By.Id("kw"));
                                    //await IWebElement q = _driver.FindElement(By.XPath("//*[@class='s_ipt']"));
                                    //Thread.Sleep(5000);
                                    //IWebElement q = _driver.FindElement(By.XPath("/*[@id='form']/span[1]"));
                                    //IWebElement q = _driver.FindElement(By.XPath("//Edit[@id='kw']"));
                                    //q.Click();
                                    //Thread.Sleep(1000);
                                    //_driver.FindElement(By.Id("su")).Click();


                                    //TOAN : 08/05/2021. search with www.so.com
                                    IWebElement q = _driver.FindElement(By.Id("input"));
                                    q.SendKeys("最新电影");
                                    Thread.Sleep(1000);
                                    _driver.FindElement(By.Id("search-button")).Click();
                                    //Thread.Sleep(3000);
                                    Thread.Sleep(1000);
                                }
                                else
                                {
                                    Thread.Sleep(3000);
                                    IWebElement q = _driver.FindElement(By.Id("query"));
                                    q.SendKeys("최신영화순위");
                                    //TOAN : 06/08/2023. button id변경
                                    //_driver.FindElement(By.Id("search_btn")).Click();
                                    _driver.FindElement(By.Id("search-btn")).Click();
                                }


                            }
                            catch(Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                            }
                            //TOAN End

                            do
                            {

                              
                                //TOAN : 06/30/2019. Background Worker Cancel확인. iterateRanking working이 아닌곳에서도
                                //발생할수 있다.
                                if (this.workerCancelCheck(e) == true)
                                {
                                    return;
                                }


                                try
                                {
                                    //TOAN : 06/30/2019. timing적으로 현재코드에 있어야지
                                    //테스트시간 출력 시 시간 Delay가 없다.
                                    //CASE1에 경우는 iterateRanking자체 코드 지연으로 인해
                                    //체크 조건까지 2분정도가 더 걸린다.

                                    //_driver.Url = _startURL;
                                    //IWebElement q = _driver.FindElement(By.Id("query"));

                                    //q.SendKeys("최신영화순위");
                                    //_driver.FindElement(By.Id("search_btn")).Click();
                                    //iterateRanking_cn
                                    //TOAN : 07/15/2021. Korea/China검증 환경 구분(China의 지역적인 한계)
                                    if (_currRegion.Equals("CN"))
                                    {
                                        System.Diagnostics.Debug.WriteLine(string.Format("TO DO: China WebSurfing"));
                                        this.iterateRanking_cn(e);
                                    }
                                    else
                                    {
                                        this.iterateRanking(e);
                                    }


                                    //this.iterateRanking(e);

                                    //TOAN : 06/15/2020. Power-control
                                    Thread.Sleep(1000);


                                }
                                catch (OpenQA.Selenium.UnhandledAlertException ex)
                                {
                                    System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                                    _driver.SwitchTo().Alert().Accept();
                                }
                                catch (OpenQA.Selenium.WebDriverException ex)
                                {
                                    System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                                }

                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                                }
                                finally
                                {
                                    System.Diagnostics.Debug.WriteLine(string.Format("Do Finally Block"));
                                }
                            } while (exit_flag == false);


                            break;
                        }
                    case "ACTION_END":
                        {

                            System.Diagnostics.Debug.WriteLine(string.Format("WEB ACTION END"));
                            break;
                        }

                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
            }
            finally
            {

            }

        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //this.TaskUpdateView(TaskStatus.TASK_RUNNING);
            //_progressform.progressBar1.Value = e.ProgressPercentage;
            //UI스레드 처리 가능

            switch (e.ProgressPercentage)
            {
                case 1:
                    {
                        this.TaskUpdateView(TaskStatus.TASK_RUNNING);
                        break;
                    }
             
                case 2:
                    {

                        break;
                    }

                default:
                    break;
            }

        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //UI스레드 처리가능
            if (e.Cancelled)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("[Web Actor]cancel completed"));
                this.TaskUpdateData(TaskStatus.TASK_FINISH);
                this.TaskUpdateView(TaskStatus.TASK_FINISH);

                //low battery로 종료되었을때는 completed가 되어도 다른 작업을 시작하면 안된다.
                if (!this._testTerminate)
                {
                    //TOAN : 07/02/2019. 
                    //Get LastTask and compare current task.
                    //If it is not last, run next task, If It is last, run first task
                    LinkedListNode<TaskRunningList> currNode = _uiManager.getTaskList().Find(TaskRunningList.TASK_WEBACTOR);
                    if (currNode != null)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("I'm WEBACTOR"));
                        LinkedListNode<TaskRunningList> nextNode = this.getNextTask(currNode);
                        _uiManager.startTask(nextNode);
                    }
                }else
                {
                    _uiManager.finishTest();
                }
            }
            else
            {
                //TO DO : What happen
                System.Diagnostics.Debug.WriteLine(string.Format("[Web Actor]work completed"));
            }
        }

        
        public bool workerCancelCheck(DoWorkEventArgs e)
        {
            bool retValue = false;

            if (worker.CancellationPending == true)
            {
                //CWebManager  자체의 system timer가 동작중이라면 종료시킨다.
                if (_systemTimer.Enabled)
                {
                    _systemTimer.Stop();
                }

                e.Cancel = true;
                exit_flag = true;
                _driver.Quit();
                retValue = true;
                //return;
            }

            return retValue;
        }

        //TOAN : 07/15/2021. SESC(CN)환경 crawling(start)
        public void iterateRanking_cn(DoWorkEventArgs e)
        {
            //China Version에 맞게 iterate ranking을 수행 한다.
            string currentXPath_part1 = "";
            string currentXPath_part2 = "";
            string currentXPath_part3 = "";
            string composeXPath = "";


            //TOAN : 08/05/2021. Below is baidu path
            //currentXPath_part1 = "//*[@id='1']/div/div/div[2]/div[1]/";
            //currentXPath_part2 = "div[1]";
            //currentXPath_part3 = "/p[1]/a/img";

            //string xpath_p1 = "div[";
            //string xpath_variable;
            //string xpath_p3 = "]";

            //TOA : 08/05/2021. Below is www.so.com path
            currentXPath_part1 = "//*[@id='mohe-relation_video_rank']/div/div[3]/div[2]/div[1]/ul/";
            currentXPath_part2 = "li[1]";
            currentXPath_part3 = "/div/a/img";

            string xpath_p1 = "li[";
            string xpath_variable;
            string xpath_p3 = "]";


            //TOAN : 05/24/2021. 아래 action은 의미가 없다.
            _saveURL = _driver.Url;

            //baidu에서는 1page당 2줄 8개까지 썸네일에 보인다. 화면에 보이게 한다.
            //element index는 1부터 시작 한다.
            for (int i = 1; i <= 8; i++)
            {
                if (this.workerCancelCheck(e) == true)
                {
                    return;
                }

                xpath_variable = i.ToString();
                currentXPath_part2 = xpath_p1 + xpath_variable + xpath_p3;
                composeXPath = currentXPath_part1 + currentXPath_part2 + currentXPath_part3;

                System.Diagnostics.Debug.WriteLine(string.Format("compose xpath : {0}", composeXPath));
                System.Diagnostics.Debug.WriteLine("[Web Actor]send find element ");
                Thread.Sleep(5000);
                //_webDriver.FindElement(By.XPath(composeXPath)).GetAttribute("value");
                _driver.FindElement(By.XPath(composeXPath)).Click();
                System.Diagnostics.Debug.WriteLine("[Web Actor]After find element ");
                Thread.Sleep(7000);

                //현재 Tab을 close시킨다.
                var tabs = _driver.WindowHandles;
                if (tabs.Count > 1)
                {
                    //Thread.Sleep(7000);
                    _driver.SwitchTo().Window(tabs[1]);
                    _driver.Close();
                    _driver.SwitchTo().Window(tabs[0]);

                    //_webDriver.Url = _saveURL;
                    //Thread.Sleep(3000);
                    //_webDriver.Navigate().Refresh();
                    //page refresh
                    //_webDriver
                    //_webDriver.get(driver.getCurrentUrl());
                }
            }
        }
        //TOAN (end)


        public void iterateRanking(DoWorkEventArgs e)
        {
            string currentXPath_part1 = "";
            string currentXPath_part2 = "";
            string currentXPath_part3 = "";
            string composeXPath = "";

            //TOAN : 09/04/2020. Naver Page소스 변경에 따른 URL변경(2.1.1.7에 포함). 그리고 인기영화순위 10개로 변경.
            //currentXPath_part1 = "//*[@class='movie_audience_ranking _main_panel v2']//div[1]//ul[1]//";
            //currentXPath_part2 = "li[1]";
            //currentXPath_part3 = "//div[1]//a[1]";


            currentXPath_part1 = "//*[@class='list_image_info type_pure_top']//div//ul[1]//";
            currentXPath_part2 = "li[1]";
            currentXPath_part3 = "//a";

            string xpath_p1 = "li[";
            string xpath_variable;
            string xpath_p3 = "]";

            _saveURL = _driver.Url;


            //TOAN : 08/22/2019 Code Change. Exception처리 후, 아래코드가 있어야지 Browser가 재동작함.
            _driver.Url = _saveURL;

            //TOAN : 09/04/2020. 인기영화순위 8->10개로 변경
            for (int i = 1; i <= /*8*/10; i++)
            {

                //TOAN : 06/30/2019. Background Worker Cancel확인
                if(this.workerCancelCheck(e)==true)
                {
                    return;
                }

                xpath_variable = i.ToString();
                currentXPath_part2 = xpath_p1 + xpath_variable + xpath_p3;
                composeXPath = currentXPath_part1 + currentXPath_part2 + currentXPath_part3;

                System.Diagnostics.Debug.WriteLine(string.Format("compose xpath : {0}",composeXPath));
                System.Diagnostics.Debug.WriteLine("[Web Actor]send find element ");
                Thread.Sleep(5000);
                _driver.FindElement(By.XPath(composeXPath)).Click();
                System.Diagnostics.Debug.WriteLine("[Web Actor]After find element ");

                Thread.Sleep(7000);
                _driver.Navigate().Back();

            }

        }

        //public void handleMovieRanking()
        //{
        //    System.Diagnostics.Debug.WriteLine(string.Format("Enter Handle MovieRanking"));
        //    string testStartTime = _myUtility.getCurrentTime();
        //    _startBattery = _myUtility.getBatteryLifeV1();
        //    exit_flag = false;

        //    //Terminate condition of below code is simply two case
        //    //First  : Testcase time meet test terminate time.
        //    //Second : Durint testing, Testcase meet test terminate battery. 

        //    do
        //    {
        //        try
        //        {
        //            _driver.Url = _startURL;
        //            IWebElement q = _driver.FindElement(By.Id("query"));

        //            q.SendKeys("최신영화순위");
        //            _driver.FindElement(By.Id("search_btn")).Click();

        //            this.iterateRanking();
        //        }
        //        catch (OpenQA.Selenium.UnhandledAlertException ex)
        //        {
        //            System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
        //            _driver.SwitchTo().Alert().Accept();
        //        }
        //        catch (OpenQA.Selenium.WebDriverException ex)
        //        {
        //            System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
        //        }

        //        catch (Exception ex)
        //        {
        //            System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
        //        }
        //        finally
        //        {
        //            System.Diagnostics.Debug.WriteLine(string.Format("Do Finally Block"));
        //        }
        //    } while (exit_flag == false);

        //    //TOAN : 06/06/2017. terminate webdriver
        //    _driver.Quit();
        //}



      
        public void setSystemTimer(int duration_sec)
        {
            //TOAN : 01/15/2019. 아래 두줄을 참조해서 진행해야 한다.
            //Dictionary<string, string> taskResult = this.composeTaskResult(TaskStatus.TASK_RUNNING);
            //_uiManager.HandleTaskReport(taskResult, TaskStatus.TASK_RUNNING);

            _systemTimer = new System.Timers.Timer();
            //_systemTimer.Interval = 5000;
            _systemTimer.Interval = duration_sec * 1000;
            _systemTimer.Elapsed += SystemTimer_Elapsed;
            _systemTimer.Start();
            //this.TaskRunningRecord(TaskRunningList.TASK_WEBACTOR);

        }

        private void SystemTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //MessageBox.Show("System Timer Expire!");
            _systemTimer.Stop();

            //TOAN : 06/30/2019. 이제 일을 그만하자.
            worker.CancelAsync();
            System.Diagnostics.Debug.WriteLine(string.Format("[Web Actor]system timer expired"));
           
        }

        public void terminateTask()
        {
            if (_systemTimer.Enabled)
            {
                _systemTimer.Stop();
            }
            //TOAN : 02/10/2019. This is loop exit flag
            lock (this)
            {
                exit_flag = true;
            }
            this.TaskUpdateData(TaskStatus.TASK_FINISH);
        }


        public void setTimer(int duration_sec)
        {
            // System.Windows.Forms.Timer formtimer = new System.Windows.Forms.Timer();
            // formtimer.Interval = duration_sec * 1000;
            // formtimer.Tick += formtimer_Tick; 

            // System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            //Initialize ListItem


            _timer = new System.Windows.Forms.Timer();
            //timer.Interval = 1000* duration_sec; // 1초
            _timer.Interval = 9000;
            _timer.Tick += new EventHandler(TimerEventProcessor);
            _timer.Start();
        }

        public void releaseTime()
        {

        }

        public void TimerEventProcessor(object sender, EventArgs e)
        {
            //timeLeft--;

            //if (timeLeft <= 0)
            //{
            //    timer.Stop();
            //    label1.Show();
            //    button1.Show();
            //}
            MessageBox.Show("Timer Expire!");
            //exit_flag = true;
            _timer.Stop();

        }

    }
}
