/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

  2019-06-30 : Add new Video Play with Window Medis Player Automation class
  2019-06-30 : Background Worker를 사용시 다음과 같이 역할 분담을 한다.
  
--***********************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using WMPLib;

using System.ComponentModel;
using System.Threading;
using OpenQA.Selenium.Remote;
using System.IO;

namespace PerformanceUsability
{

    class CVideoManager : CSeleniumBase
    {
        System.Timers.Timer _systemTimer;
        MediaPlayType _playType;
        System.Diagnostics.ProcessStartInfo _ps;
        Process _ps1;
        string _filepath;
        WMPLib.WindowsMediaPlayer Player;
        int _duration_time;

        public System.ComponentModel.BackgroundWorker worker;
        public bool _workComplete { get; set; }


        //player.PlayStateChange += new AxWMPLib._WMPOCXEvents_PlayStateChangeEventHandler(player_PlayStateChange);
        //AxWMPLib.AxWindowsMediaPlayer _player;
        public CVideoManager()
        {
            System.Diagnostics.Debug.WriteLine("Default Constructor");
            _exit_flag = false;
        }

        public CVideoManager(MediaPlayType type) : base()
        {

            System.Diagnostics.Debug.WriteLine("Real Usage Constructor");
            _playType = type;
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
                if (_systemTimer != null)
                {
                    if (_systemTimer.Enabled)
                    {
                        _systemTimer.Stop();
                    }
                }

                this.terminateVideo();
                e.Cancel = true;
                _exit_flag = true;
                retValue = true;
            }

            return retValue;
        }


        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            string argument = e.Argument as string;
            this._exit_flag = false;

                switch (argument)
                {
                case "ACTION_START":
                        {
                        
                         this.TaskUpdateData(TaskStatus.TASK_RUNNING);
                         this.TaskRunningRecord(TaskRunningList.TASK_MEDIAPLAYER);
                         worker.ReportProgress(1); //View Update
                        
                         playVideo(_filepath);

                        try
                        {
                            do
                            {
                                //worker cancel check
                                if (this.workerCancelCheck(e) == true)
                                {
                                    return;
                                }

                                //TOAN : 06/15/2020.
                                //제어권을 반환해야지 파워소모를 줄일 수가 있다.
                                Thread.Sleep(1000);
                            } while (this._exit_flag == false);

                        }catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
                        }


                        break;
                        }
                case "ACTION_END":
                       {
                           System.Diagnostics.Debug.WriteLine(string.Format("WMP ACTION END"));
                        break;
                        }
                default:
                        break;
                }
          

        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
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
                System.Diagnostics.Debug.WriteLine(string.Format("[WMP]cancel completed"));
                this.TaskUpdateData(TaskStatus.TASK_FINISH);
                this.TaskUpdateView(TaskStatus.TASK_FINISH);

                //TOAN : 07/02/2019. 
                //Get LastTask and compare current task.
                //If it is not last, run next task, If It is last, run first task
                //low battery로 종료되었을때는 completed가 되어도 다른 작업을 시작하면 안된다.
                if (!this._testTerminate)
                {
                    LinkedListNode<TaskRunningList> currNode = _uiManager.getTaskList().Find(TaskRunningList.TASK_MEDIAPLAYER);
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
                System.Diagnostics.Debug.WriteLine(string.Format("[WMP]work completed"));
            }
        }

        public void setSystemTimer(int duration_sec)
        {
            _systemTimer = new System.Timers.Timer();
            //_systemTimer.Interval = 5000;
            _systemTimer.Interval = duration_sec * 1000;
            _systemTimer.Elapsed += SystemTimer_Elapsed;
            _systemTimer.Start();

            //Dictionary<string, string> taskResult = this.composeTaskResult(TaskStatus.TASK_RUNNING);
            //_uiManager.HandleTaskReport(taskResult, TaskStatus.TASK_RUNNING);

        }

        private void SystemTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (_systemTimer != null)
            {
                if (_systemTimer.Enabled)
                {
                    _systemTimer.Stop();
                }
            }

            worker.CancelAsync();
        }


        public void setFilePath(string filename)
        {
            _filepath = filename;
        }

        public void playVideo(string filename)
        {
            //TOAN : 01/20/2019. playtime을 계산해야 한다.
            int playOffset = 7; //player가 실행되는 시간을 계산.player가 화면에 보이는 시간

            _filepath = filename;
            _duration_time = this.checkPlayTime();

            //"C:\autotest\sea sea set.avi"
            //TOAN : 02/24/2019. Below is original code
            //_ps = new System.Diagnostics.ProcessStartInfo("C:\\Program Files\\Windows Media Player\\wmplayer.exe", filename);
            //_ps.Arguments = "/fullscreen " + filename;
            //System.Diagnostics.Process.Start(_ps);

            //_duration_time += playOffset;
            //this.setSystemTimer(_duration_time);


            //TOAN : 02/24/2019. Below is new code
            //_ps = new System.Diagnostics.ProcessStartInfo("C:\\Program Files\\Windows Media Player\\wmplayer.exe", filename);
            _ps = new System.Diagnostics.ProcessStartInfo();
            _ps.FileName = "wmplayer.exe";
            _ps.Arguments = "/fullscreen " + "\"" + _filepath + "\"";
            System.Diagnostics.Process.Start(_ps);

            //_ps1 = Process.Start(System.IO.Path.Combine("C:/Program Files/Windows Media Player/wmplayer.exe", filename));
            _duration_time += playOffset;
            this.setSystemTimer(_duration_time);

        }

        //Auto Implementaion Property
        public bool _exit_flag
        {
            get;
            set;
        }

        public void terminateVideo()
        {
            //Process[] prs = Process.GetProcesses();
            //_ps1.Kill();

            //foreach (Process pr in prs)
            //{
            //    if (pr.ProcessName == "Windows Media Player")
            //    {
            //        pr.Kill();
            //    }

            //}

            //TOAN ::  01/13/2019.  commnet
            var proc = Process.GetProcessesByName("wmplayer");

            if (proc.Length > 0)
            {
                proc[proc.Length - 1].Kill();
            }

        }

        public void stopVideo()
        {
            //TimeSpan.FromSeconds(
        }

        public int checkPlayTime()
        {
            int durationTime;
            var player = new WindowsMediaPlayer();
            var clip = player.newMedia(_filepath);
            TimeSpan tDuration = TimeSpan.FromSeconds(clip.duration);
            string cduration = TimeSpan.FromSeconds(clip.duration).ToString();

            durationTime = Int32.Parse(tDuration.TotalSeconds.ToString());

            System.Diagnostics.Debug.WriteLine("total duration seconds:{0}", durationTime);
            System.Diagnostics.Debug.WriteLine("total duration:{0}", cduration);
            System.Diagnostics.Debug.WriteLine("Duration:{0}", TimeSpan.FromSeconds(clip.duration));

            //Console.WriteLine(TimeSpan.FromSeconds(clip.duration));
            return durationTime;
        }

        public void terminateTask()
        {
            if (_systemTimer != null)
            {
                if (_systemTimer.Enabled)
                {
                    _systemTimer.Stop();
                }
            }

            worker.CancelAsync();
        }


    }
}
