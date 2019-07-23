/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

  2019-06-30 : Add new File Download Manager Automation class 
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

using System.IO;
using OpenQA.Selenium.Support.UI;

using System.ComponentModel;

namespace PerformanceUsability
{
    //class CDownLoadManager
    //{
    //}

    class CDownLoadManager : CSeleniumBase
    {
        //Date member & member function 
        //public WebType _webType { get; set; }

        //Declare WebElement
        IWebElement _UserName;
        IWebElement _PassWord;
        IWebElement _loginButton;
        IWebElement _FileDownload;

        string _sUserName;
        string _sPassWordName;
        string _testURL;

        //TOAN : 12/31/2018. Data member related with Timer
        System.Timers.Timer _dTaskTimer;

        int _timerCounter;
        Boolean _targetExist;
        bool _timerComplete;
        //bool _exit_flag;
        string _downloadPath;
        string _downloadFileName;
        string _combineOperand = "\\";
        string _targetFilePath;

        //KeyList _keyList;
        //CUtility _myUtility;

        int _testCounter;
        int _testLimitation;

        string _saveURL;


        bool _isFinishTimerElapsed;

        //TOAN : 07/01/2019. Add Background Worker
        public int _finishTime;
        public System.ComponentModel.BackgroundWorker worker;
        public bool _workComplete { get; set; }

        //TOAN : 03/18/2019. Explicit Wait code
        private void MyExplicitWait()
        {
            WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromMinutes(1));

            Func<IWebDriver, bool> waitForElement = new Func<IWebDriver, bool>((IWebDriver Web) =>
            {
                Console.WriteLine(Web.FindElement(By.Id("target")).GetAttribute("innerHTML"));
                return true;
            });
            wait.Until(waitForElement);

            //ExpectedConditions형태는 deprecated가 되었다.
            //따라서 until method바로 뒤에 delegate형태로 코드를 작성해야 한다.
            //WebDriverWait wait1 = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
            //IWebElement button = wait.Until(ExpectedConditions.ElementExists(By.Id("someId"));
        }

        private string getCurrentTime()
        {
            string startTime;
            string fstartTime;

            _isFinishTimerElapsed = false;
            fstartTime = string.Format("{0:hh:mm tt}", DateTime.Now);
            startTime = System.DateTime.Now.ToString();

            startTime = System.DateTime.Now.ToString();
            System.DateTime sDisplayTime = System.Convert.ToDateTime(startTime);
            //txtStart.Text = sDisplayTime.ToString();

            return sDisplayTime.ToString();
        }

        public CDownLoadManager(WebType type) : base(type)
        {
            System.Diagnostics.Debug.WriteLine("Constructor with WebTyp  e");
            _webType = type;

            //_keyList = KeyList.Instance;
            //_myUtility = CUtility.Instance;

            //_testURL = "http://cocoknight.dothome.co.kr/frmLogin.php";
            _testURL = @"http://www.codeapril.com/frmLogin.php";
            //_saveURL = "http://www.codeapril.com/filedownloader.php";
            _sUserName = "pctest";
            _sPassWordName = "pw.1234";

            _targetExist = false;
            _timerComplete = false;
            _exit_flag = false;
            _task_exit_flag = false;
            //File Download Path 
            //TOAN : 02/10/2019. download완료 후, 삭제하기 위함.
            _downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
            //downloadFileName = "Sample_Movie.mp4";
            _downloadFileName = "sea_original.mp4";
            _targetFilePath = _downloadPath + _combineOperand + _downloadFileName;

            //TOAN : 01/28/2019. Test Limitation을 임의로 10회로 지정.
            _testCounter = 0;
            _testLimitation = 20;

           // _finishTime = 20;

            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
        }

        public void setTestTime(int time)
        {
            _finishTime = time*60;
        }

        public bool workerCancelCheck(DoWorkEventArgs e)
        {
            bool retValue = false;

            if (worker.CancellationPending == true)
            {
                //CWebManager  자체의 system timer가 동작중이라면 종료시킨다.
                if (_dTaskTimer != null)
                {
                    if (_dTaskTimer.Enabled)
                    {
                        _dTaskTimer.Stop();
                    }
                }

                _driver.Quit();
                e.Cancel = true;
                _task_exit_flag = true;
                retValue = true;
            }

            return retValue;
        }


        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            string argument = e.Argument as string;
            this._task_exit_flag = false;
            this._exit_flag = false;

            switch (argument)
            {
                case "ACTION_START":
                    {
                        this.initSelenium(0);
                        this.setTimeWait(20);
                        this.TaskUpdateData(TaskStatus.TASK_RUNNING);
                        this.TaskRunningRecord(TaskRunningList.TASK_STORAGE_ACTOR);
                        worker.ReportProgress(1);

                        this.setTaskTimer(_finishTime);

                        while (this._task_exit_flag == false)
                        {
                            try
                            {
                                //TOAN : 07/04/2019. 기존 download된것을 지우고
                                //exit flag을 false로해서 다시 waiting되도록 한다.
                                //this.cleardownloadFile();
                                this._exit_flag = false;

                                if (this.workerCancelCheck(e) == true)
                                {
                                    return;
                                }

                                //this.cleardownloadFile();
                                this.handleDownload();



                                //TOAN : 07/01/2019. wait for downloading complete
                                do
                                {
                                    //TOAN : 07/02/2019. 별도의 timer없이 loop에서 체크.
                                    if ((_timerComplete = this.IsDownLoadComplete()) == true)
                                    {
                                        this._exit_flag = true;
                                        Thread.Sleep(2000);
                                    }
                                } while (this._exit_flag == false);

                            }
                            catch (OpenQA.Selenium.WebDriverException ex)
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));

                                //_storageActor.cleardownloadFile();
                            }
                            finally
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Do Finally Block"));
                            }

                        }

                            break;
                    }
                case "ACTION_END":
                    {

                        System.Diagnostics.Debug.WriteLine(string.Format("WEB ACTION END"));
                        break;
                    }
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
                System.Diagnostics.Debug.WriteLine(string.Format("[File DownLoad]cancel completed"));
                this.TaskUpdateData(TaskStatus.TASK_FINISH);
                this.TaskUpdateView(TaskStatus.TASK_FINISH);

                //TOAN : 07/02/2019. 
                //Get LastTask and compare current task.
                //If it is not last, run next task, If It is last, run first task
                //low battery로 종료되었을때는 completed가 되어도 다른 작업을 시작하면 안된다.
                if (!this._testTerminate)
                {
                    LinkedListNode<TaskRunningList> currNode = _uiManager.getTaskList().Find(TaskRunningList.TASK_STORAGE_ACTOR);
                    if (currNode != null)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("I'm DOWNLOAD MANAGER"));
                        LinkedListNode<TaskRunningList> nextNode = this.getNextTask(currNode);
                        _uiManager.startTask(nextNode);
                    }
                }
                else
                {
                    _uiManager.finishTest();
                }

            }
            else
            {
                System.Diagnostics.Debug.WriteLine(string.Format("[File DownLoad]work completed"));
            }
        }

        public void handleDownload()
        {
            System.Diagnostics.Debug.WriteLine(string.Format("Start Downloader selenium"));
            //this.initSelenium(_webType);

            //TOAN : 03/19/2019. add page-load wait
            //_driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(500);

            //_driver.Navigate().GoToUrl(_testURL);
            _driver.Url = _testURL;
            System.Diagnostics.Debug.WriteLine("Timer Expire:{0}", this.getCurrentTime());
            Thread.Sleep(3000);

            _UserName = _driver.FindElement(By.Id("username"));
            _PassWord = _driver.FindElement(By.Id("password"));
            _loginButton = _driver.FindElement(By.Id("sbutton"));

            System.Diagnostics.Debug.WriteLine("Timer Expire:{0}", this.getCurrentTime());

            _UserName.Click();
            _UserName.Clear();
            _UserName.SendKeys(_sUserName);

            _PassWord.Click();
            _PassWord.Clear();
            _PassWord.SendKeys(_sPassWordName);

            //TOAN : 12/24/2018. Code-Debugging. component의 값을 가지고 오기
            //String uName = _UserName.GetAttribute("value");
            //String pName = _PassWord.GetAttribute("value");

            //System.Diagnostics.Debug.WriteLine("User Name:" + uName);
            //System.Diagnostics.Debug.WriteLine("Pawword:" + pName);

            _loginButton.Click();

            Thread.Sleep(3000);

            //Webpage change to filedownloader.php(expected. If need, I have to add exception control)
            _FileDownload = _driver.FindElement(By.Id("downloadLink")); ;
            _FileDownload.Click();

            //TOAN : 2018/12/28. Client에 다운로드가 완료된 이벤트를 획득한 후, 후속작업이 필요하다.
            //HTTP Protocol과 PHP언어에 대한 추가 지식이 필요한 부분이기도 하다.
            //HTTP Protocol자체는 HTTP헤더를 통한 다운로드 실행시, Client에 다운로드가 종료되었다는 이벤트는 별도로
            //주지 않는다. 따라서 W/A형태로 Client쪽에서 파일 다운로드 완료 확인을 Timer를 통해서 체크한다.
            //추가적으로 HTTP에서 Client요청없이 임의로 HTTP Response를 전송할수 있는지는 확인이 필요하다.
            //개인적인 생각으로는 안될 듯 하다.(push를 이요하지 않는 이상)

            //Add 1 sec system timer
            //_mediaPlayer.setSystemTimer(1);
            //this.setSystemTimer(1);

        }//end of function


        

        //TOAN : 04/11/2019. Add Task Task Timer
        //TOAN : 06/07/2019. enhancement for TaskTimer
        public void setTaskTimer(int duration_sec)
        {
            _task_exit_flag = false;
            _dTaskTimer = new System.Timers.Timer();
            _dTaskTimer.Interval = duration_sec * 1000;
            _dTaskTimer.Elapsed += TaskTimer_Elapsed;
            _dTaskTimer.Start();
            _isFinishTimerElapsed = false;

        }

        private void TaskTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (_dTaskTimer != null)
            {
                if (_dTaskTimer.Enabled)
                {
                    _dTaskTimer.Stop();
                }
            }

            worker.CancelAsync();
        }

       
        public void cleardownloadFile()
        {
            bool retValue = false;
            string targetFilePath = this.getTargetFilePath();

            if (retValue = File.Exists(targetFilePath))
            {
                File.Delete(targetFilePath);
            }
        }

        public void cleareUsedTimer()
        {
            if (_dTaskTimer != null)
            {
                if (_dTaskTimer.Enabled)
                {
                    _dTaskTimer.Stop();
                }
            }
            
        }

        private bool IsDownLoadComplete()
        {
            //다음 Cycle이 시작하기전에 기존 다운받았던 파일은 삭제한다.
            bool retValue = false;
            //string cpath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
            //_targetFilePath
            retValue = File.Exists(_targetFilePath);
            return retValue;
        }

        //string cpath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
        //System.Diagnostics.Debug.WriteLine(cpath);

        //        string fName = "Sample_Movie.mp4";
        //string combineOperand = "\\";
        //string fFullName = cpath + combineOperand + fName;

        //TOAN : 04/11/2019
        public bool _task_exit_flag
        {
            get;
            set;
        }
        //Auto Implementaion Property
        public bool _exit_flag
        {
            get;
            set;
        }

        public string getTargetFilePath()
        {
            return _targetFilePath;
        }

        //Downloader의 경우는 현재 루틴을 종료만 한다.
        //바로 Task Finish리포트를 하지 않는 이유는
        //web surfing처럼 부가 for loop로 인한 delay는 없기 때문이다.
        public void terminateTask()
        {
            if (_dTaskTimer != null)
            {
                if (_dTaskTimer.Enabled)
                {
                    _dTaskTimer.Stop();
                }
            }

        }

        public void refreshPage()
        {
            try
            {
                _driver.Quit(); //This is for command test
                _driver.Navigate().GoToUrl(_testURL);
            }
            catch (Exception e)
            {
                //throw (e);
                System.Diagnostics.Debug.WriteLine("[Re-Try]All Exception");
                throw (e);
            }
        }

    }
}
