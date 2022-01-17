/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

  2019-06-30 : Add new YouTube Automation Manager class
  2019-06-30 : Background Worker를 사용시 다음과 같이 역할 분담을 한다.
  2019-07-24 : Test Regison설정.

  dowork : working thread. 필요한 작업 수행
  worker_ProgressChanged : UI작업 수행. UI Task에 작업 요청 가능
  worker_RunWorkerCompleted : UI작업 수행. UI Task에 작업 요청 가능

  .Timer사용시 유의사항
  -Timer자체가 생성되지 않았는데 stop을 하면 null pointer에러가 발생
  -Null Pointer check후 실행.

  key word : how to cancel background worker after specified time in c#
  URL : https://stackoverflow.com/questions/1341488/how-to-cancel-background-worker-after-specified-time-in-c-sharp

  key word : Using Timer inside a BackGroundWorker
  URL : https://stackoverflow.com/questions/6704195/using-timer-inside-a-backgroundworker

  key word : Background Worker Cancel
  URL : https://www.wpf-tutorial.com/misc/cancelling-the-backgroundworker/
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
//using System.Timers.Timer;

using OpenQA.Selenium.Remote;
using System.Windows.Forms;
using System.IO;
using System.ComponentModel;
using System.Threading;
using OpenQA.Selenium.Remote;


namespace PerformanceUsability
{
    class CYoutubeManager : CSeleniumBase
    {
        public ControlType playerMode { get; set; }
        //public WebType _webType { get; set; }
        //IWebDriver _driver;

        KeyList _keyList;
        CUtility _myUtility;

      
        //TOAN : 06/06/2019. Set Task Finish Timer. 네트워크 연결 오류 등 Exception상황 발생 시
        //Youtube영상 자체는 30분이지만 상황에 따라 더 많은 시간이 걸릴때도 있다.
        //따라서 You-tube task를 finish하기 위해서, 매 1초마다 체크하는 systemtimer와 finishtimer를 같이 고려
        //finishtimer가 발동하기 전에 systemtimer에 의해서 youtube가 종료되었다면, finishtimer를 무시하고,
        //finishtimer로 확인시점까지 systemtimer에 의해 youtube영상이 play end되어지지 않았으면, youtube task는
        //강제종료 한다.
        System.Timers.Timer _finishTimer;
        bool _isFinishTimerElapsed;
        bool _isSkipAdvertisement;
        int _timerCounter;
        Boolean _targetExist;
        bool _isVideoEnd;

        //protected Form1 _uiManager;

        //TOAN : 06/30/2019. add background worker
        public System.ComponentModel.BackgroundWorker worker;
        public bool _workComplete { get; set; }
        public string _playURL;
        public int _finishTime ;

        //TOAN : 07/24/2019. Test Region설정
        public string _testRegion;


        //TOAN : 06/25/2021. retry counter
        public int _youtube_retry_count = 0;
        public CYoutubeManager()
        {
            //System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}",currObj.Key,currObj.Value);
            System.Diagnostics.Debug.WriteLine("Default Constructor for CMediaPlay");
            System.Diagnostics.Debug.WriteLine("Player Default Mode:{0}", this.playerMode);

            _testRegion = "ALL";
            // this.initSelenium();
        }

        //  public CWebActor(WebType type):base(type) 
        public CYoutubeManager(WebType type) : base(type)
        {
            System.Diagnostics.Debug.WriteLine("Constructor with WebType");
            _webType = type;
            _targetExist = false;
            _isVideoEnd = false;
            _isFinishTimerElapsed = false;
            _isSkipAdvertisement = false;
            //this.initSelenium(type);
            _testRegion = "ALL";

            _keyList = KeyList.Instance;
            _myUtility = CUtility.Instance;

            _exit_flag = false;
            _finishTime = 120; //2분. 실제 Release할 때는 혹시 모르니 30분으로 걸어놓는다.

            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
        }

        public void setURL(string url)
        {
            _playURL = url;
        }

        public void checkAdvertise()
        {
            By checkTarget = By.XPath("//button[@class='ytp-ad-skip-button ytp-button']");

            if (_targetExist = this.IsElementPresent(checkTarget) == true)
                {
                    System.Diagnostics.Debug.WriteLine("Find Skip AdverTise Attribute");
                    //_driver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='HD'])[1]/following::span[1]")).Click();
                    _driver.FindElement(checkTarget).Click();
                    _isSkipAdvertisement = true;
                }
           
        }

       

        public void checkvideoEnd()
        {
            //If we check video end about normal case
            string sCurrentValue;
            sCurrentValue = _driver.FindElement(By.XPath("//*[@id='movie_player']")).GetAttribute("class");
            _isVideoEnd = this.IsVideoEnded(sCurrentValue);
            //System.Diagnostics.Debug.WriteLine("Player State :", sCurrentValue);

            if (_isVideoEnd == true)
            {
                worker.CancelAsync();
            }
            
        }

        public void setTestTime(int time)
        {
            _finishTime = time * 60;
        }

        public void setRegion(string region)
        {
            _testRegion = region;

        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
           string argument = e.Argument as string;
           this._exit_flag = false;

            try
            {
                switch (argument)
                {
                    case "ACTION_START":
                        {
                            //TOAN : 07/15/2021. code-change
                            //this.initSelenium(0);
                            this.initSelenium(_webType);

                            //TOAN : 07/16/2021. Exception Handling. Exception이 생기더라도 Timer에 의해 종료되도록 지원
                            this.setTaskFinishTimer(_finishTime);
                            try
                            {
                                this.setTimeWait(5);
                                this.TaskUpdateData(TaskStatus.TASK_RUNNING);
                                this.TaskRunningRecord(TaskRunningList.TASK_YOUTUBE);
                            }catch(Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                            }
                            //TOAN End.

                            worker.ReportProgress(1); //View Update(This is very nice code)

                            this.playVideoStreaming(_playURL);

                            //TOAN : 07/16/2021. PCAUT처럼 Simple하게 변경
                            ////TOAN : 07/24/2019.
                            ////_uiManager
                            //string currRegion = _uiManager.getCurrentRegion();

                            //if (!currRegion.Equals("CN"))
                            //{
                            //    this.controlVideoStreaming(ControlType.FULLSCREEN);
                            //}
                            ////this.controlVideoStreaming(ControlType.FULLSCREEN);


                            //Thread.Sleep(5000);
                            //TOAN End.

                            //Thread.Sleep(2000);
                            //this.setSystemTimer(1);

                            //this.setTaskFinishTimer(_finishTime);

                            try
                            {
                                do
                                {
                                    //worker cancel check
                                    if (this.workerCancelCheck(e) == true)
                                    {
                                        return;
                                    }

                                    //TOAN : 07/16/2021. PCAUT와 동일하게 변경(simple)
                                    ////step1 : check advertisement
                                    ////youtube영상은 광고 있는게 있고, 없는것도 있다.(다시보기 했을때)
                                    ////이경우 advertistmemt check를 하지 않으면, 원치않게 광고가 끝났을 때, 광고를 test contents로 알고 종료함
                                    ////TOAN : 07/24/2019. SESC QQ Player는 youtube와 구조가 틀리기 때문에
                                    ////Dependency가 있는 코드는 사용하지 않는다.
                                    //if (!currRegion.Equals("CN"))
                                    //{
                                    //    if (_isSkipAdvertisement == false)
                                    //    {
                                    //        this.checkAdvertise();

                                    //    }
                                    //}

                                    ////step2 : check video end
                                    ////TOAN : 07/02/2019. 별도의 timer없이 loop에서 체크.
                                    ////TOAN : 07/24/2019. SESC QQ Player는 youtube와 구조가 틀리기 때문에
                                    ////Dependency가 있는 코드는 사용하지 않는다.
                                    //if (!currRegion.Equals("CN"))
                                    //{
                                    //    this.checkvideoEnd();
                                    //}

                                    //TOAN : 06/15/2020.
                                    //TOAN (End)

                                    Thread.Sleep(1000);
                                } while (this._exit_flag == false);
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
                System.Diagnostics.Debug.WriteLine(string.Format("[Youtube Play]cancel completed"));
                this.TaskUpdateData(TaskStatus.TASK_FINISH);
                this.TaskUpdateView(TaskStatus.TASK_FINISH);

                //TOAN : 07/02/2019. 
                //Get LastTask and compare current task.
                //If it is not last, run next task, If It is last, run first task
                //low battery로 종료되었을때는 completed가 되어도 다른 작업을 시작하면 안된다.
                if (!this._testTerminate)
                {
                    LinkedListNode<TaskRunningList> currNode = _uiManager.getTaskList().Find(TaskRunningList.TASK_YOUTUBE);
                    if (currNode != null)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("I'm YOUTUBE"));
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
                System.Diagnostics.Debug.WriteLine(string.Format("[Youtube Play]work completed"));
            }
        }


        public bool workerCancelCheck(DoWorkEventArgs e)
        {
            bool retValue = false;

            if (worker.CancellationPending == true)
            {
                //CWebManager  자체의 system timer가 동작중이라면 종료시킨다.
                if (_finishTimer != null)
                {
                    if (_finishTimer.Enabled)
                    {
                        _finishTimer.Stop();
                    }
                }
                e.Cancel = true;
                _exit_flag = true;
                _driver.Quit();
                retValue = true;
            }

            return retValue;
        }
        //public void initSelenium(WebType mode)
        //{
        //    //IWebDriver driver = new ChromeDriver();
        //    //driver.Url = "https://www.youtube.com/watch?v=WhSGqlqyXq0";

        //    switch (mode)
        //    {
        //        case WebType.WEB_Chrome:
        //            {
        //                _driver = new ChromeDriver();
        //                break;
        //            }
        //        case WebType.WEB_FireFox:
        //            {
        //                _driver = new FirefoxDriver();
        //                break;
        //            }
        //        case WebType.WEB_IE:
        //            {
        //                _driver = new InternetExplorerDriver();
        //                break;
        //            }
        //        case WebType.WEB_EDGE:
        //            {
        //                _driver = new EdgeDriver();
        //                break;
        //            }
        //        default:
        //            break;
        //    }
        //}

        public void playVideoStreaming(string url)
        {
            //IWebDriver driver = new ChromeDriver(); 
            //driver.Url = "https://www.youtube.com/watch?v=WhSGqlqyXq0";
            //this.initSelenium(url);

            //_driver = new ChromeDriver();
            //_driver.Url = "https://www.youtube.com/watch?v=WhSGqlqyXq0";



            //TOAN : 11/04/2021. PCAUT style - start
            bool b_exception_fire = false;

            do
            {
                //TOAN : 07/15/2021.
                b_exception_fire = false;
                try
                {
                    _driver.Manage().Window.Maximize();
                    _driver.Url = _playURL;
                }
                //catch (OpenQA.Selenium.WebDriverException ex)
                //{
                //    System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                //}

                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                    //TOAN : 07/15/2021. code 보강
                    b_exception_fire = true;
                    _youtube_retry_count = _youtube_retry_count + 1;
                }

                //_youtube_retry_count = _youtube_retry_count + 1;
                Thread.Sleep(2000);

            } while (b_exception_fire == true && _youtube_retry_count < 3);

            _youtube_retry_count = 0;

            try
            {

                ////*[@id="movie_player"]/div[5]/button
                //IWebElement q = _webDriver.FindElement(By.Id("query"));
                //"//button[@class='ytp-ad-skip-button ytp-button']"

                //TOAN : 01/07/2021. 아래 xpath에 해당하는 target은 youtube player의 중간에 나타나는 play버튼이다.
                IWebElement play_button = _driver.FindElement(By.XPath("//*[@id=\"movie_player\"]/div[5]/button"));
                play_button.Click();
                System.Diagnostics.Debug.WriteLine(string.Format("Movie player overlay playbutton clicked"));

            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
            }
            //TOAN : 11/04/2021. PCAUT style - end




            //TOAN : 11/04/2021. original code-start
            //            try
            //            {
            //                //TOAN : 07/16/2021. PCAUT내용 적용
            //                _driver.Manage().Window.Maximize();
            //                _driver.Url = url;
            //                //TOAN : 08/23/2018. Full Screen버튼에 click이벤트 적용 후, 다시 창모드로 전환이 된다.
            //                //Video Streaming이 출력되고 난후에, Full Screen버튼을 누르면 이 현상이 없어진다.
            //                //즉, 해당 element가 존재하지 않아서 발생한 이슈는 아니다.
            //                //우선 Sleep Command로 이슈가 개선되는지 확인해 보자. sleep command로 증상 개선을 확인 함.

            //                _isSkipAdvertisement = false;
            //                _targetExist = false;dowor
            //                _isVideoEnd = false;
            //                _isFinishTimerElapsed = false;
            //                _isSkipAdvertisement = false;
            //                _exit_flag = false;

            //Thread.Sleep(6000);
            //                //Application.DoEvents();
            //            }
            //            catch (Exception ex)
            //            {
            //                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
            //                throw ex;
            //            }
            //            finally
            //            {

            //            }

            //TOAN : 11/04/2021. original cod end
        }

        public void controlVideoStreaming(ControlType mode)
        {
            switch (mode)
            {
                case ControlType.READY:
                    {
                        break;
                    }

                case ControlType.PLAY:
                    {
                        break;
                    }
                case ControlType.SKIP:
                    {
                        break;
                    }
                case ControlType.STOP:
                    {
                        break;
                    }
                case ControlType.VOLUME:
                    {
                        break;
                    }
                case ControlType.SETTING:
                    {
                        //HD 720p설정.
                        //WebElement ele = driver.findElements(By.xpath("your xpath"));
                        // WebDriverWait wait = new WebDriverWait(driver, 50);
                        // wait.until(ExpectedConditions.elementToBeClickable(ele));

                        _driver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='실시간'])[1]/following::button[2]")).Click();
                        _driver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='품질'])[1]/following::span[1]")).Click();
                        _driver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='HD'])[1]/following::span[1]")).Click();

                        //_driver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='실시간'])[1]/following::button[2]")).Click();
                        //_driver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='품질'])[1]/following::div[2]")).Click();
                        //_driver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='HD'])[1]/following::span[1]")).Click();
                        break;
                    }
                case ControlType.FULLSCREEN:
                    {
                        //System.Diagnostics.Debug.WriteLine("FullScreen Control");
                        //TOAN : 08/22/2018. Below code cause Compound class names not permitted error WebDriver 런타임 에러 발생
                        //IWebElement fullScreen = _driver.FindElement(By.ClassName("ytp-fullscreen-button ytp-button")); 

                        //TOAN : 08/23/2018. Below code is correct
                        //IWebElement fullScreen = _driver.FindElement(By.ClassName("ytp-fullscreen-button"));
                        //IWebElement fullScreen = _driver.FindElement(By.CssSelector(".ytp-fullscreen-button.ytp-button"));
                        IWebElement fullScreen = _driver.FindElement(By.XPath("//button[@class='ytp-fullscreen-button ytp-button']"));
                        fullScreen.Click();

                        //TOAN : 08/22/2018. Firefox katalon recorder에서 추출한 명령어 형태. 
                        //"//button[@class='ytp-large-play-button ytp-button']"
                        //_driver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='실시간'])[1]/following::button[6]")).Click();
                        break;
                    }


                default:
                    break;
            }
        }

        public void playVideoLocal()
        {

        }

        public void closeSelenium()
        {

        }

        //TOAN : 06/06/2019. Set Task Terminate Time. In case of youtube is 30 minutes
        public void setTaskFinishTimer(int duration_sec)
        {
            //_finishTimer
            _finishTimer = new System.Timers.Timer();
            _finishTimer.Interval = duration_sec * 1000;
            _finishTimer.Elapsed += FinishTimer_Elapsed;
            _finishTimer.Start();
        }

        private void FinishTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (_finishTimer != null)
            {
                if (_finishTimer.Enabled)
                {
                    _finishTimer.Stop();
                }
            }
                //_isFinishTimerElapsed = true;
            
            //TOAN : 06/30/2019. 이제 일을 그만하자.
            worker.CancelAsync();
        }

      
        private bool IsVideoEnded(string chkString)
        {
            bool retValue = false;
            string cmpString = "ended-mode";

            retValue = chkString.Contains(cmpString);
            return retValue;
        }
        private bool IsElementPresent(By by)

        {
            try
            {
                //Exception이 안생기면 By로 지정된 target이 존재한다는 것이다.
                _driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        //Auto Implementaion Property
        public bool _exit_flag
        {
            get;
            set;
        }

        public void terminateTask()
        {
            //terminate task가 호출되었다는 말은 task상태가running된것이다.
            if (_finishTimer != null)
            {
                if (_finishTimer.Enabled)
                {
                    _finishTimer.Stop();
                }
            }
            //_isFinishTimerElapsed = true;

            //TOAN : 06/30/2019. 이제 일을 그만하자.
            worker.CancelAsync();
            
        }
        
    }



}