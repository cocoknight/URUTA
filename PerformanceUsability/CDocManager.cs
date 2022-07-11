
/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

  2019-06-30 : Add new PPT Automation Manager class 
  2022-01-17 : Power Point 실행 후 종료 루틴 추가 (W//A)
  - process kill로 power point 앱 종료진행
--***********************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//TOAN : 08/30/2018. Related with Document Automation with using PowerPoint or Excel
using System.Reflection;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;


using System.Threading;
using System.ComponentModel;
using System.Diagnostics;

namespace PerformanceUsability
{
    //class CDocManager
    //{
    //}
    class CDocManager : CSeleniumBase
    {
        public DocType docMode { get; set; }
        System.Timers.Timer _dTaskTimer;

        //TOAN : 08/31/2018. 객체에 접근 제한 속성을 두지 않는다면. Default접근 제한은 private속성이다.
        //C#의 Data-Member는 선언과 동시에 초기화가 필요 없지만, 지역변수의 경우는 선언과함께 초기화기 필요하다.
        //그렇지 않으면 run-time error가 발생 한다.

        Application _pptApplication;
        Presentation _pptPresentation;
        CustomLayout _customLayout;

        Slides _slides;
        _Slide _slide;
        TextRange _objText;

        
        int _docPageCount = 0;

        //TOAN : 07/01/2019. Add Background Worker
        public int _finishTime;
        public System.ComponentModel.BackgroundWorker worker;
        public bool _workComplete { get; set; }

        public CDocManager(WebType type) : base(type)
        {

            _keyList = KeyList.Instance;
            _myUtility = CUtility.Instance;

            //System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}",currObj.Key,currObj.Value);
            System.Diagnostics.Debug.WriteLine("Default Constructor for CDocMaker");
            // this.initSelenium();
            _exit_flag = false;


            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += new DoWorkEventHandler(worker_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(worker_ProgressChanged);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
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

                //TOAN : 01/17/2022. Power-Point Application을 종료 시킨다.
                this.terminate_ppt();
                //_driver.Quit();
                 e.Cancel = true;
                _exit_flag = true;
                retValue = true;
            }

            return retValue;
        }

        public void terminate_ppt()
        {
            //var proc = Process.GetProcessesByName("Video.UI");

            //if (proc.Length > 0)
            //{
            //    proc[proc.Length - 1].Kill();
            //}

            //TOAN : 01/17/2022. 아래 코드에서 Quit가 동작하지 않는다.
            //if(_pptApplication!=null)
            //{
            //    System.Diagnostics.Debug.WriteLine(string.Format("PPT Application Terminate"));
            //   _pptApplication.Quit();

            //}

            //TOAN : 01/17/2022. Process Kill로 코드 변경.
            try
            {
                var proc = Process.GetProcessesByName("POWERPNT");
                int process_num = proc.Length;
                if (proc.Length > 0)
                {
                    do
                    {
                        proc[process_num - 1].Kill();
                        //proc = Process.GetProcessesByName("Excel");
                        process_num = process_num - 1;
                    } while (process_num > 0);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
            }

        }


        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            string argument = e.Argument as string;
            _exit_flag = false;

            switch (argument)
            {
                case "ACTION_START":
                    {
                        
                        this.TaskUpdateData(TaskStatus.TASK_RUNNING);
                        this.TaskRunningRecord(TaskRunningList.TASK_DOCUMENT);
                        worker.ReportProgress(1);


                        this.setTaskTimer(_finishTime);
                        this.initPPT();

                        int pageNum = 1;
                        try
                        {

                            do
                            {
                                if (this.workerCancelCheck(e) == true)
                                {
                                  
                                    return;
                                }

                                this.addPage(pageNum);
                                //pageNum++;
                                Thread.Sleep(2000);
                            } while (this._exit_flag == false);
                        }catch(Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                        }

                        break;
                    }
                case "ACTION_END":
                    {
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
                System.Diagnostics.Debug.WriteLine(string.Format("[PPT Manager]cancel completed"));
                this.TaskUpdateData(TaskStatus.TASK_FINISH);
                this.TaskUpdateView(TaskStatus.TASK_FINISH);


                //TOAN : 07/02/2019. 
                //Get LastTask and compare current task.
                //If it is not last, run next task, If It is last, run first task
                //low battery로 종료되었을때는 completed가 되어도 다른 작업을 시작하면 안된다.
                if (!this._testTerminate)
                {
                    LinkedListNode<TaskRunningList> currNode = _uiManager.getTaskList().Find(TaskRunningList.TASK_DOCUMENT);
                    if (currNode != null)
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("I'm PPT MAKER"));
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
                System.Diagnostics.Debug.WriteLine(string.Format("[PPT Manager]work completed"));
            }
        }


        public void setTestTime(int time)
        {
            _finishTime = time * 60;
        }

        //TOAN : 06/08/2019. Code Enhancement
        public void addPage(int num)
        {
            //TOAN : 06/08/2019. Add-Page을 실행할 때, 아래 2줄은 꼭 쌍으로 참조가 되어야 한다.
            //TOAN : 07/11/2022. PowerPoint가 중복으로 Access할때 예외발생하는 이슈 처리
            //e.g] 유효기간이 지났을 때 Log-in화면 처리등의 예외 추가.
            try
            {

                _slides = _pptPresentation.Slides;
                _slide = _slides.AddSlide(/*1*/num, _customLayout);

                _objText = _slide.Shapes[1].TextFrame.TextRange;
                _objText.Text = "제목입니당";
                _objText = _slide.Shapes[2].TextFrame.TextRange;
                //_objText.Text = "1번째줄\n2번째줄\n3번째줄";

                int ioop = 0;
                for (ioop = 0; ioop < 10; ioop++)
                {
                    _objText.Text += ioop.ToString();
                    _objText.Text += "\n";
                    Thread.Sleep(2000);
                }

                _slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "여기는 슬라이드 설명쓰는곳입니당.";
            }catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                this.terminate_ppt();
                this.initPPT();
            }
            //TOAN END : 07/11/2022
        }

        public void addPageWithTime()
        {
            int i = 1;
            _slides = _pptPresentation.Slides;
            do
            {
                if (_exit_flag == true)
                    break;

                _slide = _slides.AddSlide(i, _customLayout);
                // 타이틀 추가
                _objText = _slide.Shapes[1].TextFrame.TextRange;
                _objText.Text = "제목입니당";
                //TOAN : 01/28/2019. FontName에서 Exception이 발생하는듯 하다.
                //_objText.Font.Name = "Gulim";
                //_objText.Font.Size = 32;
                _objText.Font.Size = 20;
                _objText = _slide.Shapes[2].TextFrame.TextRange;
                //_objText.Text = "1번째줄\n2번째줄\n3번째줄";
                int ioop = 0;
                //int ioop = 48;
                for (ioop = 0; ioop < 10; ioop++)
                {
                    _objText.Text += ioop.ToString();
                    _objText.Text += "\n";
                    //아래 코드를 사용하면 루프중에 프로그램이 종료되버린다.
                    Delay(1000);
                }

                //TOAN : 01/28/2019. Temporary Blocking
                _slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "여기는 슬라이드 설명쓰는곳입니당.";

                i = i + 1;
            } while (_exit_flag == false);
        }

        public void addPageContent(int pageNum)
        {
            _slides = _pptPresentation.Slides;

            for (int i = 1; i <= pageNum; i++)
            {
                if (_exit_flag == true)
                    break;
                //_slides = _pptPresentation.Slides;
                //_slide = _slides.AddSlide(1, _customLayout);
                _slide = _slides.AddSlide(i, _customLayout);
                // 타이틀 추가
                _objText = _slide.Shapes[1].TextFrame.TextRange;
                _objText.Text = "제목입니당";
                //TOAN : 01/28/2019. FontName에서 Exception이 발생하는듯 하다.
                //_objText.Font.Name = "Gulim";
                //_objText.Font.Size = 32;
                _objText.Font.Size = 20;
                _objText = _slide.Shapes[2].TextFrame.TextRange;
                //_objText.Text = "1번째줄\n2번째줄\n3번째줄";
                int ioop = 0;
                //int ioop = 48;
                for (ioop = 0; ioop < 10; ioop++)
                {
                    _objText.Text += ioop.ToString();
                    _objText.Text += "\n";
                    //아래 코드를 사용하면 루프중에 프로그램이 종료되버린다.
                    //Delay(1000);
                }

                //TOAN : 01/28/2019. Temporary Blocking
                _slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "여기는 슬라이드 설명쓰는곳입니당.";
            }
        }

        public void initPPT()
        {
            //_pptApplication = new Application();
            _pptApplication = new PowerPoint.Application();

            _pptPresentation = _pptApplication.Presentations.Add(Office.MsoTriState.msoTrue);

            _customLayout = _pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

            //_pptApplication.Visible = true;

            //_pptApplication.Quit();

            
        }

        public void initPPTAutomation()
        {
            _pptApplication = new Application();

            _pptPresentation = _pptApplication.Presentations.Add(Office.MsoTriState.msoTrue);

            _customLayout = _pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

            // 새 슬라이드 생성
            _slides = _pptPresentation.Slides;
            _slide = _slides.AddSlide(1, _customLayout);

            // 타이틀 추가
            _objText = _slide.Shapes[1].TextFrame.TextRange;
            _objText.Text = "제목입니당";
            //TOAN : 01/28/2019. FontName에서 Exception이 발생하는듯 하다.
            //_objText.Font.Name = "Gulim";
            //_objText.Font.Size = 32;
            _objText.Font.Size = 20;
            _objText = _slide.Shapes[2].TextFrame.TextRange;
            //_objText.Text = "1번째줄\n2번째줄\n3번째줄";
            int ioop = 0;
            //int ioop = 48;
            for (ioop = 0; ioop < 10; ioop++)
            {
                _objText.Text += ioop.ToString();
                _objText.Text += "\n";
                //아래 코드를 사용하면 루프중에 프로그램이 종료되버린다.
                //Delay(1000);
            }

            //TOAN : 01/28/2019. Temporary Blocking
            _slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "여기는 슬라이드 설명쓰는곳입니당.";

            //File저장 영역은 현재 디렉토리에 저장을 시켜 준다.
            //_pptPresentation.SaveAs(@"c:\COOLA\ppttest.pptx", PpSaveAsFileType.ppSaveAsDefault, Office.MsoTriState.msoTrue);

            _slide = _slides.AddSlide(2, _customLayout);
            // 타이틀 추가
            _objText = _slide.Shapes[1].TextFrame.TextRange;
            //_objText.Text = "제목입니당";
            _objText.Font.Size = 32;
            _objText.Text = "제목";

            //Delay(5000);

            _objText.Text = _objText.Text + "입니당";
            _objText.Font.Name = "Gulim";
            _objText.Font.Size = 32;

            _objText = _slide.Shapes[2].TextFrame.TextRange;
            _objText.Text = "4번째줄\n5번째줄\n6번째줄";

            _slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "여기는 슬라이드 설명쓰는곳입니당.";


        }

        public void actionDocMaker(DocType mode)
        {
            switch (mode)
            {
                case DocType.DOC_READY:
                    {
                        break;
                    }
                case DocType.DOC_TABLE:
                    {
                        break;
                    }
                case DocType.DOC_PICTURE:
                    {
                        break;
                    }
                case DocType.DOC_SHAPE:
                    {
                        break;
                    }
                case DocType.DOC_CHART:
                    {
                        break;
                    }
                case DocType.DOC_TYPING:
                    {

                        break;
                    }

                default:
                    break;
            }
        }

        private static DateTime Delay(int MS)
        {
            DateTime ThisMoment = DateTime.Now;
            TimeSpan duration = new TimeSpan(0, 0, 0, 0, MS);
            DateTime AfterWards = ThisMoment.Add(duration);

            while (AfterWards >= ThisMoment)
            {
                System.Windows.Forms.Application.DoEvents();
                ThisMoment = DateTime.Now;
            }

            return DateTime.Now;
        }

        public void setTaskTimer(int duration_sec)
        {
            _dTaskTimer = new System.Timers.Timer();
            _dTaskTimer.Interval = duration_sec * 1000;
            _dTaskTimer.Elapsed += TaskTimer_Elapsed;
            _dTaskTimer.Start();
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


        //Auto Implementaion Property
        public bool _exit_flag
        {
            get;
            set;
        }

    }

}
