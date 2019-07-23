/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

  2019-06-30 : Add new Web automation Manager class
  2019-06-30 : 

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
                            this.initSelenium(0);
                            this.setTimeWait(100);
                            this.TaskUpdateData(TaskStatus.TASK_RUNNING);
                            this.TaskRunningRecord(TaskRunningList.TASK_WEBACTOR);
                            worker.ReportProgress(1); //View Update

                            //this.setSystemTimer(/*600*/_finishTime);
                            exit_flag = false;


                            //1time query
                            _driver.Url = _startURL;
                            IWebElement q = _driver.FindElement(By.Id("query"));

                            q.SendKeys("최신영화순위");
                            _driver.FindElement(By.Id("search_btn")).Click();


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

                                    this.iterateRanking(e);

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

        public void iterateRanking(DoWorkEventArgs e)
        {
            string currentXPath_part1 = "";
            string currentXPath_part2 = "";
            string currentXPath_part3 = "";
            string composeXPath = "";

            currentXPath_part1 = "//*[@class='movie_audience_ranking _main_panel v2']//div[1]//ul[1]//";
            currentXPath_part2 = "li[1]";
            currentXPath_part3 = "//div[1]//a[1]";

            string xpath_p1 = "li[";
            string xpath_variable;
            string xpath_p3 = "]";

            _saveURL = _driver.Url;

            for (int i = 1; i <= 8; i++)
            {

                //TOAN : 06/30/2019. Background Worker Cancel확인
                if(this.workerCancelCheck(e)==true)
                {
                    return;
                }

                xpath_variable = i.ToString();
                currentXPath_part2 = xpath_p1 + xpath_variable + xpath_p3;
                composeXPath = currentXPath_part1 + currentXPath_part2 + currentXPath_part3;

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
