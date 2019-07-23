/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

    2019-01-02 : Display user interface and Handling User menu event

    2019-01-28 : modify HandleFiledownload
                 Add a multiple download code with small file.
                 Current Web Server doesn't surrpot large size file download.(e.g over 1GB)
    
    2019-02-05 : If Application meets low battery state, Terminate Running Task and make a report
    2019-02-17 : Sequence Change From Task Start to Task End for avoiding ui update issue. This is task synchroize issue 
    2019-02-25 : File Download종료 후 재시작관련 에러 수정 (1.0.0.6 fix)
        - HandleFileDownLoad()의 _downLoadNumber 및 txtPassNumber.Text초기화
    2019-03-31 : Add CReportMaker for Report Handling
    2019-04-04 : Add Runnint time columns for Listview
    2019-04-05 : CTestDecision Instance Add
    2019-04-07 : Call MakeDecision method
    2019-04-11 : change Default Running Time
    2019-04-11 : Add HandleFileDownloadV1 
    2019-05-12 : Modify WebActor Speed 
    2019-05-12 : Modify Exception Handling fow Web Actor with code integration
    2019-05-13 : Code Refactoring(HandleMovieRanking)
    2019-06-06 : Before test running, Check each task checkbox
    2019-06-10  : Add playvideostreaming exception handling
    2019-06-25 : Add web drive quit code befort starting current task
    2019-07-02 : check test order. (checkTestOrder)
--***********************************************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//TOAN : 08/30/2018. Related with Document Automation with using PowerPoint or Excel
using System.Reflection;

using Office = Microsoft.Office.Core; 
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;


//TOAN : 10/14/2018. Process Start, Like as Media Player
using System.Diagnostics;

//TOAN : 12/12/2018. Get Power Information with Battery
using System.Management;
using System.Threading;
using System.IO;
using Microsoft.Office.Interop.Excel;

//TOAN : 03/20/2019. network check
using System.Net.NetworkInformation;

//TOAN : 04/08/2019. 글자만 추출하기 위함.
using System.Text.RegularExpressions;



namespace PerformanceUsability
{
    public partial class Form1 : Form
    {
        public ControlType _playerMode { get; set; }
        public WebType _webType { get; set; }
        public MediaPlayType _mediaPlayerType { get; set; }
        public TaskRunningList _taskRunningList { get; set; }

       
        
        CReportMaker _reportMaker;
        //TOAN : 06/30/2019. add new web manager
        CWebManager _webManager;

        //TOAN : 06/30/2019. add new youtube manager
        CYoutubeManager _youtubeManager;

        //TOAN : 06/30/2019. add new WMP Manager
        CVideoManager _videoManager;

        //TOAN : 06/30/2019. add new file download Manager
        CDownLoadManager _downloadManager;

        //TOAN : 06/30/2019. add new PPT Manager
        CDocManager _docManager;

        //TOAN : 04/05/2019
        CTestDecision _testDecision;

        int _downLoadLimitation;
        int _downLoadNumber;
        KeyList _keyList;
        CUtility _myUtility;

        int _testcase_no;
        //TOAN : 12/17/2018. Declare Generic Data Structure for ListView
        public Dictionary<string, string> _columnInfoDic;
        //public List<Dictionary<string, object>> _testcaseList; 
        public List<Dictionary<string,string>> _iteminfoList;

        //TOAN : 01/13/2018. WMP(Window Media Player)Control
        System.Diagnostics.ProcessStartInfo _ps;

        //TOAN : 01/13/2018. WMP TestCode
        WMPLib.WindowsMediaPlayer _Player;

        string _loadURL;

        //TOAN : 01/31/2019. Task Terminate
        public bool _bTerminate;
        System.Timers.Timer _systemTimer;

        //TOAN : 03/20/2019. add monitor timer
        System.Timers.Timer _monitorTimer;
        string _runningTask;
        int _lowBattery;

        //TOAN : 06/10/2019. add youtube retry counter
        public int _youtubeRetry;
        public bool _youtube_exceptionHappen;
        public int _youtubeFinishTime = 30;
        //initialize string
        //enter time in minutes
        //string sDownloadNumber = "Enter Download Number(50-100)";
        //TOAN : 04/11/2019. Download횟수에서 시간컨셉으로 변경.(횟수로 했을때 전류 소모량이 틀려짐)
        string sDownloadNumber = "Enter time in minutes"; 
        string sPassNumber = "Pass Number";
        string sWorkingTime = "Enter time in minutes";
        //string sWorkingTime = "Enter the Web Surfing Time";
        //string sPageNumber = "Enter the number of pages you want to create.";
        //string sPageNumber = "Enter the Document Editing Time";
        string sPageNumber = "Enter time in minutes";

        //TOAN : 07/02/2019. TOAN LinkedList사용(Doubly Linked List)
        //LinkedList<int> _tasklist;
        //TaskRunningList
        LinkedList<TaskRunningList> _tasklist;
        List<TaskRunningList> _taskOrder;
        public Form1()
        {

            //TOAN : 12/24/2018 . Avoid System.InvalidOperationException: '크로스 스레드 작업이 잘못되었습니다. 
            //위 Exception을 해결하는 가장 Simple한 방법
            //아니면 Thread Racing에 맞게 코드 변경해야 한다.
            CheckForIllegalCrossThreadCalls = false;

            InitializeComponent();
            //TOAN : 12/24/2018. Display Thread Number
            
            //initialize ListView Data Structure
            _columnInfoDic = new Dictionary<string, string>();
            _iteminfoList = new List<Dictionary<string, string>>();

            //Singletone instance를 먼저 생성한다.
            //Using C# Property. Below is singletone property
            _keyList = KeyList.Instance;
            _myUtility = CUtility.Instance;

            //int batterySize = Int32.Parse(txtBattery.Text);
            //byte batterySize = byte.Parse(txtBattery.Text);
            //string currBattery = txtBattery.Text;
            //int batterySize = Int32.Parse(currBattery);
            //_myUtility.setBatteryWH(batterySize);

            //setup class module
            //TOAN : 03/27/2019. Web Type Test with IE
            _playerMode = ControlType.READY;
            _webType = WebType.WEB_Chrome;



            //TOAN : 06/30/2019. add WebManager
            _webManager = new CWebManager(_webType);
            _webManager.testcase_no = "1";
            _webManager.testcase_name = lblCase1.Text;
            _webManager.connectUI(this);

            //TOAN : 06/30/2019. add Youtube Manager
            _youtubeManager = new CYoutubeManager(_webType);
            _youtubeManager.testcase_no = "2";
            _youtubeManager.testcase_name = lblCase2.Text;
            _youtubeManager.connectUI(this);

            //TOAN : 06/30/2019. add WMP manager
            _mediaPlayerType = MediaPlayType.MEDIA_WMP;
            _videoManager = new CVideoManager(_mediaPlayerType);
            _videoManager.testcase_no = "3";
            _videoManager.testcase_name = lblCase3.Text;
            _videoManager.connectUI(this);

            //TOAN : 06/30/2019. add File DownLoad Manager
            _downloadManager = new CDownLoadManager(_webType);
            _downloadManager.connectUI(this);
            _downloadManager.testcase_no = "4";
            _downloadManager.testcase_name = lblCase7.Text;

            //_docManager
            //TOAN : 06/30/2019. add new PPT Manager
            _docManager = new CDocManager(_webType);
            _docManager.connectUI(this);
            _docManager.testcase_no = "5";
            _docManager.testcase_name = lblCase4.Text;

           
            //string currTC = _keyList.getTC();
            string currTC = _keyList.k_testcase;
            //string currTC = KeyList.k_testcase;
            //System.Diagnostics.Debug.WriteLine("Key Command :{0} ",currTC);
            //MessageBox.Show("Current Test case : " + currTC);
            //System.Diagnostics.Debug.WriteLine("Player Default Mode:{ 0}",this.playerMode);
            //_mediaPlayer = new CMediaPlay();

            _bTerminate = false;
            this.initializeListColumns();

            //TOAN : 01/31/2019
            //_runningTask = "empty";

            //TOAN : 03/31/2019. instance for ReportMaker
            _reportMaker = new CReportMaker();
            _reportMaker.connectUI(this);

            //TOAN : 04/05/2019. instance for TestDecision
            _testDecision = new CTestDecision();
            _testDecision.connectUI(this);

            //TOAN : 06/10/2019
            _youtubeRetry = 0;
            _youtube_exceptionHappen = false;

            //TOAN : 07/02/2019.
            _tasklist = new LinkedList<TaskRunningList>();
            _taskOrder = new List<TaskRunningList>();

            //order input
            _taskOrder.Add(TaskRunningList.TASK_WEBACTOR);
            _taskOrder.Add(TaskRunningList.TASK_YOUTUBE);
            _taskOrder.Add(TaskRunningList.TASK_MEDIAPLAYER);
            _taskOrder.Add(TaskRunningList.TASK_STORAGE_ACTOR);
            _taskOrder.Add(TaskRunningList.TASK_DOCUMENT);

            //LinkedList 값 넣기. _tasklist is not circular list
            //for (TaskRunningList i = TaskRunningList.TASK_WEBACTOR; i <=TaskRunningList.TASK_DOCUMENT; i++)
            //{
            //    _tasklist.AddLast(i);
            //}

            //LinkedList내용 출력하기

            //foreach (var item in _tasklist)
            //{
            //    //System.Console.WriteLine($"{item}");
            //    System.Diagnostics.Debug.WriteLine(string.Format("node value:{0}", item));
            //}

            //get last node and print next node
            //LinkedListNode<TaskRunningList> node = _tasklist.Last;

          
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Form Load");
            //TOAN : 03/20/2919. Add Event Handler
            NetworkChange.NetworkAvailabilityChanged += new NetworkAvailabilityChangedEventHandler(NetworkChange_NetworkAvailabilityChanged);

            //initialize textbox
            txtTime1.Text = sWorkingTime;
            txtDownloadTime.Text = sDownloadNumber;
            //txtPassNumber.Text = sPassNumber;
            txtPPTWorkingTime.Text = sPageNumber;

            //this.setCheckBox_All_Value(1);
           // this.enableCheckBox_All_Value(0);

            this.checkDefault();

            //TOAN : 07/02/2019. Task별 Default Time을 보이게 ㅎㄴ다.
            this.checkDefaultTime();

            //TOAN : 07/02/2019. 모든 Test를 선택하게 한다.
            this.checkDefaultTest();

            //WMP코드 테스트
            //PlayFile(@"c:\myaudio.wma");
            //"C:\\movie\\jwonwochi\\전우치.jeonwoochi.HDTV.XviD.AC3-5.1ch-GO.avi"
            // PlayFile(@"c:\movie\jwonwochi\전우치.jeonwoochi.HDTV.XviD.AC3-5.1ch-GO.avi");

            //int batterySize = Int32.Parse(txtBattery.Text);
            //byte batterySize = byte.Parse(txtBattery.Text);
            //string currBattery = txtBattery.Text;
            //int batterySize = Int32.Parse(currBattery);
            //_myUtility.setBatteryWH(batterySize);

        }

        //TOAN : 03/20/2019. checking for network available
        //TOAN : 06/07/2019. Ignore Handler Action.
        //각 Actor의 자체 Exception처리 코드에서 에러 처리토록 변경

        private void NetworkChange_NetworkAvailabilityChanged(object sender, NetworkAvailabilityEventArgs e)
        {
            //if (e.IsAvailable)
            //{
            //    //MessageBox.Show("Available");
            //    //TOAN:03/27/2019. Temporary Blocking
            //    //driver.navigate.refresh
            //    //_webActor.refreshPage();

            //    //TOAN : 04/12/2019. 현재 Task에 따라 구분할 것
            //    switch (_taskRunningList)
            //    {
            //        case TaskRunningList.TASK_WEBACTOR:
            //            {
            //                _webActor.refreshPage();
            //                break;
            //            }
            //        case TaskRunningList.TASK_STORAGE_ACTOR:
            //            {
            //                //_testURL
            //                _storageActor.refreshPage();
            //                break;
            //            }
            //    }
            //}
            //else
            //{
            //    //MessageBox.Show("Not available");
            //    //_webActor.refreshPage();
            //}
        }

        private void groupSetting_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void lblCase4_Click(object sender, EventArgs e)
        {

        }

        private void chkCase4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cboCase4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void initializeListColumns()
        {
            RunningList.View = View.Details;
            RunningList.GridLines = true;
            RunningList.FullRowSelect = true;
            RunningList.CheckBoxes = true;

            //Add column header
            //RunningList.Columns.Add("TestCase", 300);
            //RunningList.Columns.Add("Status", 80);
            //RunningList.Columns.Add("Remaing Battery", 90);
            //RunningList.Columns.Add("Task Discharge", 90);
            //RunningList.Columns.Add("Task Discharge(wh)", 110);
            //RunningList.Columns.Add("Power Consumption", 110);
            //RunningList.Columns.Add("Start Time", 75);
            //RunningList.Columns.Add("End Time", 80);

            //RunningList.Columns.Add("TestCase", 300);


            //RunningList.Columns.Add("k_test_case", "TestCase", 300);
            //RunningList.Columns.Add("k_status","Status", 80);
            //RunningList.Columns.Add("Remaing Battery", 90);
            //RunningList.Columns.Add("Task Discharge", 90);
            //RunningList.Columns.Add("Task Discharge(wh)", 110);
            //RunningList.Columns.Add("Power Consumption", 110);
            //RunningList.Columns.Add("Start Time", 75);
            //RunningList.Columns.Add("End Time", 80);


            //TOAN : 01/09/2018. 컬럼추가시 Key를 이용한 방법으로 변경
            RunningList.Columns.Add(_keyList.k_testcase, "TestCase", /*300*/250);
            RunningList.Columns.Add(_keyList.k_status, "Status", 80);
            RunningList.Columns.Add(_keyList.k_remaining_battery, "Remaing Battery", 90);
            RunningList.Columns.Add(_keyList.k_discharge, "Task Discharge", 90);
            RunningList.Columns.Add(_keyList.k_discharge_wh, "Task Discharge(wh)",110);
            RunningList.Columns.Add(_keyList.k_power_consumption_wh, "Power Consumption", 110);
            RunningList.Columns.Add(_keyList.k_start_time, "Start Time", 75);
            RunningList.Columns.Add(_keyList.k_end_time, "End Time", 80);
            //TOAN : 04/04/2019
            RunningList.Columns.Add(_keyList.k_running_time, "Running Time", 80);

            //TOAN : 01/09/2019. 컬럼을 추가할 때 키값을 부여한 경우
            //ContainKey 메서드를 이용하면, 해당 ListView의 컬럼이 그 키를 포함하고 있는지 유무를 확인할 수 있다.
            bool b_hasKey;
            int i_keyIndex;

            b_hasKey = RunningList.Columns.ContainsKey("k_test_case");
            i_keyIndex = RunningList.Columns.IndexOfKey("k_status");
            
            System.Diagnostics.Debug.WriteLine("Key Exist:{0}", b_hasKey);
            System.Diagnostics.Debug.WriteLine("Key Index:{0}", i_keyIndex);

            //Check number of columns
            String msg = String.Format("컬럼수=={0}", RunningList.Columns.Count);
            System.Diagnostics.Debug.WriteLine(msg);


            //Refernce Code
            //string a="aa";
            //string b="bb";
            //string c="cc";

            //TOAN : 12/24/2018. 최초 ListViewItem을 추가할때, 해당 컬럼에 해당하는 값이 없어도, ""처럼 초기값을 지정해주어야 한다.
            //이렇게 해야지 추후 해당 List가 update될때, 빈 컬럼에 대한 값을 update가능 하다.
            //즉, 전체 컬럼 갯수 만큼 초기화를 진행해야 한다.
            //string[] row = { "Naver인기 영화 순위검색", "진행중!!", "100%","","","","","" };
            ////string[] row = { a, b, c };
            //ListViewItem cItem = new ListViewItem(row);
            //RunningList.Items.Add(cItem);
            
        }  

      

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void chkCase6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void lblCase3_Click(object sender, EventArgs e)
        {

        }

        private void lblCase7_Click(object sender, EventArgs e)
        {

        }

        private void lblCase1_Click(object sender, EventArgs e)
        {

        }

        private void cboCase7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void checkDefaultTest()
        {
            //TOAN : 07/02/2019. Form이 Loading될 때, Default로 모든 Task를 선택한다.
            chkMovieRank.Checked = true;
            chkYoutube.Checked = true;
            chkLocalPlayer.Checked = true;
            chkWebDownload.Checked = true;
            chkPowerPoint.Checked = true;
        }
        private void checkDefaultTime()
        {
            //임의로 코드 테스트를 진행 한다.

            //txtTime1.Text = "5";
            //txtDownloadNumber.Text = "10";
            //txtPPTWorkingTime.Text = "5";

            //문자열비교
            System.Diagnostics.Debug.WriteLine("Text Value:{0}",txtTime1.Text);

  
            if (txtTime1.Text.Equals(sWorkingTime))
            {
                //TOAN : 04/11/2019. Change Default Time
                txtTime1.Text = /*"5"*/"20";
                txtTime1.Update();
            }
            //sDownloadNumber
            if (txtDownloadTime.Text.Equals(sDownloadNumber))
            {
                txtDownloadTime.Text = "20";
                txtDownloadTime.Update();
            }

            if(txtPPTWorkingTime.Text.Equals(sPageNumber))
            {
                txtPPTWorkingTime.Text = "30"/*"10"*/;
                txtPPTWorkingTime.Update();
            }
        }

        private void checkDefault()
        {
            //string.IsNullOrEmpty(myString)
            //txtTime1.Text

            //String fileToOpen = "C:\\movie\\Street-19627.avi";
            //String fileToOpen = txtTime3.Text;

            if(string.IsNullOrEmpty(txtTime3.Text))
            {
                //txtTime3.Text = "C:\\movie\\Street-19627.avi";
                txtTime3.Text = "C:\\autotest\\split.avi";
            }

            if (string.IsNullOrEmpty(txtTime2.Text))
            {
                //txtTime2.Text = "https://www.youtube.com/watch?v=eKNLp1xjdzI";
                txtTime2.Text = "https://www.youtube.com/watch?v=IVWeOQA9lAc"; //대도서관 30분짜리 컨텐츠(실제 검증용). release용
                //txtTime2.Text = "https://www.youtube.com/watch?v=MBNQgq56egk";   //test용

                //string playURL = "https://www.youtube.com/watch?v=WhSGqlqyXq0";
                //string playURL = "https://www.youtube.com/watch?v=PYgyTtclSxQ";
                //string playURL = "https://www.youtube.com/watch?v=IKZEmLvYVF0"; //빅뱅 러브송
                //string playURL = "https://www.youtube.com/watch?v=MBNQgq56egk"; //빅뱅 맨정신
                //string playURL = "https://www.youtube.com/watch?v=DopcD8gJW5Q"; //짧은 테스트용 youtube음악.
            }
        }

        
        private void setCheckBox_All_Value(int check)
        {

            if(check==1)
            {
                //select all checkbox
                chkMovieRank.Checked = true;
                chkYoutube.Checked = true;
                chkLocalPlayer.Checked = true;
                chkWebDownload.Checked = true;
                chkPowerPoint.Checked = true;
            }
            else if(check==0)
            {
                //unselect all checkbox 
                chkMovieRank.Checked = false;
                chkYoutube.Checked = false;
                chkLocalPlayer.Checked = false;
                chkWebDownload.Checked = false;
                chkPowerPoint.Checked = false;
            }
        }

        private void enableCheckBox_All_Value(int check)
        {

            if(check==1)
            {
                //enable check box
                chkMovieRank.Enabled = true;
                chkYoutube.Enabled = true;
                chkLocalPlayer.Enabled = true;
                chkWebDownload.Enabled = true;
                chkPowerPoint.Enabled = true;
            }
            else if(check==0)
            {
                //disable check box
                chkMovieRank.Enabled = false;
                chkYoutube.Enabled = false;
                chkLocalPlayer.Enabled = false;
                chkWebDownload.Enabled = false;
                chkPowerPoint.Enabled = false;
            }
        }

        public bool isTaskCheckable()
        {
            bool isCheck = false;

            if( (chkMovieRank.Checked == false)&&(chkYoutube.Checked==false) &&(chkLocalPlayer.Checked==false) 
                &&(chkWebDownload.Checked==false) &&(chkPowerPoint.Checked==false) )
            {
                isCheck = false;
            }
            else
            {
                isCheck = true;
            }

                    return isCheck;
        }

       
        //checkbox값을 모두 해제 시킨다.
        public void releaseCheckBox()
        {

            chkMovieRank.Checked = false;
            chkYoutube.Checked = false;
            chkLocalPlayer.Checked = false;
            chkWebDownload.Checked = false;
            chkPowerPoint.Checked = false;
        }

        //TOAN : 07/02/2019. check test order.
 
        public void checkTestOrder()
        {
            //현재 선택된 Task가 무엇인지 확인한다. 그리고 수행순서를 정한다.

            //chkMovieRank.Checked = true;
            //chkYoutube.Checked = true;
            //chkLocalPlayer.Checked = true;
            //chkWebDownload.Checked = true;
            //chkPowerPoint.Checked = true;

            foreach (TaskRunningList task in _taskOrder)
            {

                switch (task)
                {
                    case TaskRunningList.TASK_WEBACTOR:
                        {
                            //_tasklist.AddLast(i);
                            if(chkMovieRank.Checked)
                            {
                                //동일node중복 삽입 불가
                                if(_tasklist.Find(TaskRunningList.TASK_WEBACTOR)==null)
                                {
                                    _tasklist.AddLast(TaskRunningList.TASK_WEBACTOR);
                                }
                               
                            }
                            else
                            {
                                //TO DO:
                                if (_tasklist.Find(TaskRunningList.TASK_WEBACTOR) != null)
                                {
                                    _tasklist.Remove(TaskRunningList.TASK_WEBACTOR);
                                }
                            }
                            break;
                        }
                    case TaskRunningList.TASK_YOUTUBE:
                        {
                            if (chkYoutube.Checked)
                            {
                                if (_tasklist.Find(TaskRunningList.TASK_YOUTUBE) == null)
                                {
                                    _tasklist.AddLast(TaskRunningList.TASK_YOUTUBE);
                                }
                            }
                            else
                            {
                                if (_tasklist.Find(TaskRunningList.TASK_YOUTUBE) != null)
                                {
                                    _tasklist.Remove(TaskRunningList.TASK_YOUTUBE);
                                }
                            }

                            break;
                        }
                    case TaskRunningList.TASK_MEDIAPLAYER:
                        {
                            if (chkLocalPlayer.Checked)
                            {
                                if (_tasklist.Find(TaskRunningList.TASK_MEDIAPLAYER) == null)
                                {
                                    _tasklist.AddLast(TaskRunningList.TASK_MEDIAPLAYER);
                                }
                            }
                            else
                            {
                                if (_tasklist.Find(TaskRunningList.TASK_MEDIAPLAYER) != null)
                                {
                                    _tasklist.Remove(TaskRunningList.TASK_MEDIAPLAYER);
                                }
                            }

                            break;
                        }
                    case TaskRunningList.TASK_STORAGE_ACTOR:
                        {
                            if (chkWebDownload.Checked)
                            {
                                if (_tasklist.Find(TaskRunningList.TASK_STORAGE_ACTOR) == null)
                                {
                                    _tasklist.AddLast(TaskRunningList.TASK_STORAGE_ACTOR);
                                }
                            }
                            else
                            {
                                if (_tasklist.Find(TaskRunningList.TASK_STORAGE_ACTOR) != null)
                                {
                                    _tasklist.Remove(TaskRunningList.TASK_STORAGE_ACTOR);
                                }
                            }

                            break;
                        }
                    case TaskRunningList.TASK_DOCUMENT:
                        {
                            if (chkPowerPoint.Checked)
                            {
                                if (_tasklist.Find(TaskRunningList.TASK_DOCUMENT) == null)
                                {
                                    _tasklist.AddLast(TaskRunningList.TASK_DOCUMENT);
                                }
                            }
                            else
                            {
                                if (_tasklist.Find(TaskRunningList.TASK_DOCUMENT) != null)
                                {
                                    _tasklist.Remove(TaskRunningList.TASK_DOCUMENT);
                                }
                            }

                            break;
                        }
                    default:
                            break;
                }
            }

        }

        public LinkedList<TaskRunningList> getTaskList()
        {
            return _tasklist;
        }

        public void startTask(LinkedListNode<TaskRunningList> node)
        {
            //_tasklis에서 첫번째을 가지고 와서 start한다.
            //LinkedListNode<TaskRunningList> node = _tasklist.First;


            switch (node.Value)
            {

                case TaskRunningList.TASK_WEBACTOR:
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Start Web Surfing"));
                        string requestMode = "ACTION_START";
                        int testtime = Int32.Parse(txtTime1.Text);
                        _webManager.setTestTime(testtime);
                        _webManager.worker.RunWorkerAsync(requestMode);

                        break;
                    }
                case TaskRunningList.TASK_YOUTUBE:
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Start YouTube"));
                        string requestMode = "ACTION_START";
                        // string playURL = "https://www.youtube.com/watch?v=MBNQgq56egk";
                        //string playURL = "https://www.youtube.com/watch?v=PYgyTtclSxQ";  //youtube영상. teamj
                        //string playURL = "https://www.youtube.com/watch?v=IVWeOQA9lAc";    //대도서관영상.
                        string playURL = txtTime2.Text;
                        //youtube는 사용자로부터 시간입력받는것이 아니라, 고정으로 30분으로fix한다. 
                        //실제 Youtube컨텐츠의 종료를 체크할수도 있지만,이경우는 streaming상황에 따라 수행시간이 틀려질수 있다.
                        //따라서 Youtube의 경우 contents상관없이 30분 task-timer로 종료시킨다.
                        //대도서관의 경우도, 31분쯤에 contents종료가 되겠지만, task finish timer가 먼저 수행될 것이다.
                        int testtime = 30; 
                        _youtubeManager.setURL(playURL);
                        _youtubeManager.setTestTime(testtime);
                        _youtubeManager.worker.RunWorkerAsync(requestMode);

                        break;
                    }
                case TaskRunningList.TASK_MEDIAPLAYER:
                    {
                        //별도의 testtime setting이 필요 없다.
                        System.Diagnostics.Debug.WriteLine(string.Format("Start MediaPlayer"));
                        string requestMode = "ACTION_START";
                        string filePath = txtTime3.Text;
                        //string filePath = "C:\\autotest\\sea.avi";
                        _videoManager.setFilePath(filePath);
                        _videoManager.worker.RunWorkerAsync(requestMode);

                        break;
                    }
                case TaskRunningList.TASK_STORAGE_ACTOR:
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Start File Download"));
                        string requestMode = "ACTION_START";
                        //_downloadManager.setTestTime(3);
                        int testtime = Int32.Parse(txtDownloadTime.Text);
                        _downloadManager.setTestTime(testtime);
                        _downloadManager.worker.RunWorkerAsync(requestMode);

                        break;
                    }
                case TaskRunningList.TASK_DOCUMENT:
                    {
                        System.Diagnostics.Debug.WriteLine(string.Format("Start PowerPoint"));
                        string requestMode = "ACTION_START";
                        int testtime = Int32.Parse(txtPPTWorkingTime.Text);
                        _docManager.setTestTime(testtime);
                        _docManager.worker.RunWorkerAsync(requestMode);
                        break;
                    }
                default:
                    break;
            }
        }

        private void cmdRun_Click(object sender, EventArgs e)
        {
            //TOAN : 02/07/2019. Battery용량값(wh)이 입력되어 있지 않으면. 실행하지 않는다.
            //this.checkDefault();
           
            if (string.IsNullOrEmpty(txtBattery.Text)| string.IsNullOrEmpty(txtLowBattery.Text))
            {
                //Display waringing message
                if (string.IsNullOrEmpty(txtBattery.Text))
                {
                    MessageBox.Show("Battery Capacity is nesscary." + Environment.NewLine + "Input Battery Capacity.",
                                       "Message",
                                       MessageBoxButtons.OK,
                                       MessageBoxIcon.Information);
                    txtBattery.Select();
                }

                //move focus to txtBattery
                if(string.IsNullOrEmpty(txtLowBattery.Text))
                {
                    MessageBox.Show("Low Battery Level is nesscary." + Environment.NewLine + "Input Low Battery Level(%).",
                                "Message",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                    txtLowBattery.Select();
                }
               
            }else if(this.isTaskCheckable()==false)
            {
                MessageBox.Show("Testcase Selection is nesscary." + Environment.NewLine + "Select Testcase",
                                     "Message",
                                     MessageBoxButtons.OK,
                                     MessageBoxIcon.Information);
            }
            else
            {
                txtStart.Text = this.getCurrentTime();
                txtCurrentBattery.Text = getBatteryLevel();

                txtStart.Update();
                txtCurrentBattery.Update();
                //this.checkDefaultTime();


                //Test Model배터리 용량지정. 사용자가 set버튼을 안누르고 테스트시작 가능
                int batterySize = Int32.Parse(txtBattery.Text);
                _myUtility.setBatteryWH(batterySize);

                //Low Battery용량 지정. 사용자가 set버튼을 안누르고 테스트 시작 가능
                _lowBattery = Int32.Parse(txtLowBattery.Text);

                //TOAN : 07/02/2019. check test order
                this.checkTestOrder();

                //Get FirstTask and Start
                LinkedListNode<TaskRunningList> node = _tasklist.First;
                this.startTask(node);

                //register battery checktimer for every 2sec
                this.setMonitorTimer(2);
                //this.setMonitorTimer(1200);
            }

        }

        


        private void txtCurrentBattery_TextChanged(object sender, EventArgs e)
        {

        }

        private void bGetBattery_Click(object sender, EventArgs e)
        {
            //TOAN : 12/12/2018. Get current Battery Level
            String batterystatus;

            PowerStatus pwr = SystemInformation.PowerStatus;
            batterystatus = SystemInformation.PowerStatus.BatteryChargeStatus.ToString();

            MessageBox.Show("battery charge status : " + batterystatus);

            string batterylife;
            //int calBattery;
            float calBattery;

            batterylife = SystemInformation.PowerStatus.BatteryLifePercent.ToString();
            //calBattery = Int32.Parse(batterylife);
            calBattery = float.Parse(batterylife)*100;
            MessageBox.Show(calBattery.ToString());

            txtCurrentBattery.Text = calBattery.ToString();
        }

        //TOAN : 12/12/2018. Get BatteryLife
        private string getBatteryLevel()
        {
            string batterylife;
            string retValue;
            float calBattery;

            batterylife = SystemInformation.PowerStatus.BatteryLifePercent.ToString();
            calBattery = float.Parse(batterylife) * 100;
            retValue = calBattery.ToString();


            return retValue;
        }

        private string getCurrentTime()
        {
            string startTime;
            string fstartTime;

            fstartTime = string.Format("{0:hh:mm tt}", DateTime.Now);
            startTime = System.DateTime.Now.ToString();

            startTime = System.DateTime.Now.ToString();
            System.DateTime sDisplayTime = System.Convert.ToDateTime(startTime);
            //txtStart.Text = sDisplayTime.ToString();

            return sDisplayTime.ToString();
        }
        private void bGetTime_Click(object sender, EventArgs e)
        {
            //Get Current Time
            //txtStart.Text = "a";
            string startTime;
            string fstartTime;

            //fstartTime = string.Format("{0:hh:mm:ss tt}", DateTime.Now);
            fstartTime = string.Format("{0:hh:mm tt}", DateTime.Now);

            System.Diagnostics.Debug.WriteLine(fstartTime);
            //startTime = string.Format("{0:hh:mm tt}", DateTime.Now);
            //startTime = string.Format("{0:hh:mm}", DateTime.Now);  //분까지만 표시되는 Format
            //MessageBox.Show(startTime);
            startTime = System.DateTime.Now.ToString();
            System.DateTime sDisplayTime = System.Convert.ToDateTime(startTime);
            txtStart.Text = sDisplayTime.ToString();


            string currentTime = System.DateTime.Now.ToString("h:mm:ss.ff");
            //System.Diagnostics.Debug.WriteLine(currentTime);

            string startTime1;
            string currentTime1;

            //TOAN : 01/11/2019. ToString에 날짜 포맷이 있으면 오전,오후가 바뀌어서 표현된다.
            //startTime1 = System.DateTime.Now.ToString("h:mm:ss.ff");
            //currentTime1 = System.DateTime.Now.ToString("h:mm:ss.ff");

            startTime1 = System.DateTime.Now.ToString();
            currentTime1 = System.DateTime.Now.ToString();

            System.DateTime start = System.Convert.ToDateTime(startTime1);
            System.DateTime current = System.Convert.ToDateTime(currentTime1);
            System.Diagnostics.Debug.WriteLine(current.Subtract(start).TotalMinutes);


            System.TimeSpan timeCal = current - start;
            string calTime = timeCal.ToString();
            string calTime1 = string.Format("{0:hh:mm}", calTime);
            System.Diagnostics.Debug.WriteLine(calTime1);
            System.Diagnostics.Debug.WriteLine(calTime);

            //시스템 코드 경과시간 임의로 테스트하기
            // DateTime a = new DateTime(2010, 05, 12, 13, 15, 37);
            ////DateTime a = new DateTime(2010, 05, 12, 12, 15, 00);
            ////DateTime a = new DateTime(2010, 05, 11, 13, 15, 00);
            //DateTime b = new DateTime(2010, 05, 12, 13, 45, 00);
            //System.Diagnostics.Debug.WriteLine(b.Subtract(a).TotalMinutes);
            //double calTotalMninutes = b.Subtract(a).TotalMinutes;
            //double calToHour = calTotalMninutes / 60;
            ////double convertHour = Math.Round(calToHour, 2); //소숫점 2째자리
            //double convertHour = Math.Round(calToHour, 3);   //소수점 3째자리에서 반올림, 2째자리까지만 유효하게 한다.
            //System.Diagnostics.Debug.WriteLine(convertHour);

            DateTime a = System.DateTime.Now;
            
            string aaa = string.Format("{0:hh:mm:ss tt}", DateTime.Now);
            DateTime b = System.DateTime.Now;
            System.Diagnostics.Debug.WriteLine(b.Subtract(a).TotalMinutes);
            System.Diagnostics.Debug.WriteLine(a.ToString());
        }

        private void txtStart_TextChanged(object sender, EventArgs e)
        {
            //ListItem을 동적으로 Update하기
        }

        private void bAddList_Click(object sender, EventArgs e)
        {
            //ListItem을 동적으로 Insert하기 
            //URL : http://www.csharpstudy.com/WinForms/WinForms-listview.aspx
            //URL : http://blog.naver.com/PostView.nhn?blogId=curlicu&logNo=40139908594

            string[] row = {"Naver인기 영화 순위검색","진행중", "100%","100%","11wh", "77wh", "11:00","12:00"};
            //Case 1
            //ListViewItem cItem = new ListViewItem("Naver인기 영화순위 검색");
            //ListViewItem cItem = new ListViewItem("Naver인기 영화순위 검색");
            //cItem.SubItems.Add("진행중");
            //cItem.SubItems.Add("100%");

            //Case 2
            ListViewItem cItem = new ListViewItem(row); 
            RunningList.Items.Add(cItem);
          
        }
         
        private void bUpdateList_Click(object sender, EventArgs e)
        {
            //TOAN : 01/29/2019. 아래 두줄이 있어야지 List
            RunningList.Focus();
            RunningList.Select();

            //RunningList.Items[0].Focused = true;
            //RunningList.Items[0].Selected = true;


            RunningList.Items[1].Focused = true;
            RunningList.Items[1].Selected = true;

            ListViewItem currItem = RunningList.SelectedItems[0];
            string chkValue1 = currItem.SubItems[0].Text;

            //System.Diagnostics.Debug.WriteLine("{0}",chkValue1);
            System.Diagnostics.Debug.WriteLine(chkValue1);

            //CASE1 : 기존 아이템을 모두 제거하고 새롭게 등록하는 코드
            //RunningList.Items.Remove(currItem);
            //ListViewItem cItem = new ListViewItem("Naver인기 영화순위 검색");
            //cItem.SubItems.Add("완료");
            //cItem.SubItems.Add("95%"); 
            //cItem.SubItems.Add("5%");
            //cItem.SubItems.Add("3.8wh");
            //cItem.SubItems.Add("7.5w");

            //RunningList.Items.Add(cItem);

            //CASE2 : Item중 변경이 생긴 SubItem만 등록 시키는 코드
            //currItem.SubItems[1].Text="완료";
            //System.Diagnostics.Debug.WriteLine(currItem.SubItems[1.GetHashCode()].Text);
            System.Diagnostics.Debug.WriteLine(currItem.SubItems[1].GetHashCode().ToString());
            currItem.SubItems[1].Text = "완료";
            currItem.SubItems[2].Text = "95%";
            //currItem.SubItems[3].Text = "7wh";

            //RunningList.SelectedItems.Clear();
            //RunningList.Items[_testcase_no].Focused = false;
           // RunningList.Items[_testcase_no].Selected = false;
        }

        private void bDeleteList_Click(object sender, EventArgs e)
        {
            //선택된 List를 제거한다.
            //그럼, 사용자에게 보이지는 않지만, 등록할때 각 Listview의 구분 id를 부연한 후 등록한다.
            //그리고 ListView내용을 별도파일로 저장시켜야 한다.
            
            //ListViewItem currItem = RunningList.SelectedItems[0];
            //RunningList.Items.Remove(currItem );
            
            RunningList.Focus();
            RunningList.Select();
            RunningList.Items[0].Focused = true;
            RunningList.Items[0].Selected = true;

            ListViewItem currItem = RunningList.SelectedItems[0];
            RunningList.Items.Remove(currItem);

        }

        public Dictionary<string, string> getColumninfo()
        {
            Dictionary<string, string> columnInfo;
            columnInfo = _columnInfoDic;
            return columnInfo;
        }

        public void updateList(Dictionary<string, string> cList)
        {

        }

        public void HandleTaskReport(Dictionary<string, string> cList, TaskStatus status)
        {
            
            switch (status)
            {
                case TaskStatus.TASK_RUNNING:
                    {
                        //TOAN : 04/04/2019. 컬럼갯수 만큼 초기화가 필요하다.현재 8개지만 아이템이 추가될 경우
                        //stringp[] row값이 update되어야 한다.
                        //out of bound exception이 발생하지 않는 경우.
                        //컬럼갯수 만큼 subitem을 초기화 해주어야지, 향후 update시 해당 컬럼인덱스로 제어가능하다.
                        //subitem컬럼을 등록하지 않은상태에서, index로 접근하면 outof bound exception이 발생한다.
                        string[] row = { cList[_keyList.k_testcase], "", "", "", "", "", "","","" };
                        ListViewItem cItem = new ListViewItem(row);

                        RunningList.Items.Add(cItem);

                        int listsize = RunningList.Items.Count;
                        System.Diagnostics.Debug.WriteLine("ListView Count:{0}", listsize);
                     
                        //Add-operation수행 후, 선택하고 index를 찾는다.
                        //추가된것을 지우는 상황이 발생하지 않으므로 추가된후, '전체갯수'-1을 한것이
                        //현재 ListviewItem의 index로 간주할 수 있다.

                        //int i_testcase_no = Int32.Parse(cList[_keyList.k_testcase_no]) - 1;
                        //int i_testcase_no = listsize - 1;
                        _testcase_no = listsize - 1;
                        RunningList.Focus();
                        RunningList.Select();
                        
                        RunningList.Items[_testcase_no].Focused = true;
                        RunningList.Items[_testcase_no].Selected = true;

                        //SelectedItems는 선택되어진 집합을 가져오는데, 현재 선택은 하나밖에 없으므로
                        //SelectedItems[0]형태로 고정 시킨다.
                        ListViewItem currItem = RunningList.SelectedItems[/*_testcase_no*/0];


                        //TOAN : 01/19/2019. Temporary Blocking
                        this.updateListView(currItem, cList);

                        //int columnIndex = RunningList.Columns.IndexOfKey(currObj.Key);
                        //currItem.SubItems[0].Text = cList[_keyList.k_testcase];
                        //currItem.SubItems[1].Text = "Running";

                        RunningList.SelectedItems.Clear();
                        RunningList.Items[_testcase_no].Focused = false;
                        RunningList.Items[_testcase_no].Selected = false;
                        System.Diagnostics.Debug.WriteLine("TASK RUNNING");
                        //UI를 바로 update하자
                        //RunningList.Update();
                        RunningList.Refresh();
                        break;
                    }

                case TaskStatus.TASK_FINISH:
                    {
                        //update columns with task result
                        RunningList.Focus();
                        RunningList.Select();

                        RunningList.Items[_testcase_no].Focused = true;
                        RunningList.Items[_testcase_no].Selected = true;

                        ListViewItem currItem = RunningList.SelectedItems[0];

                        //currItem.SubItems[0].Text = cList[_keyList.k_testcase]; 
                        //currItem.SubItems[1].Text = "Finish";

                        //TOAN : 01/19/2019. Temporary Blocking
                        this.updateListView(currItem, cList);
                        System.Diagnostics.Debug.WriteLine("TASK FINISH");

                        RunningList.SelectedItems.Clear();
                        RunningList.Items[_testcase_no].Focused = false;
                        RunningList.Items[_testcase_no].Selected = false;
                        //UI를 바로 업데이트 하자
                        //RunningList.Update();
                        RunningList.Refresh();
                        break;
                    }
                default:
                    break;
            }

        }

        public void addListViewItem(ListViewItem cListItem, 
                                    Dictionary<string, string> updateDic)
        {
            foreach (var currObj in updateDic)
            {

                //k_testcase_no키는 Listview컬럼에 포함되지 않으므로, 해당 키의 컬럼인덱스는 존재하지 않는다.
                //따라서 해당 키의 경우는 skip하고, 나머지 키는 값을 가지고 와서 ListView에 뿌려 준다.
                if (!currObj.Key.Equals(_keyList.k_testcase_no))
                {
                    System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                    //i_keyIndex = RunningList.Columns.IndexOfKey("k_status");
                    int columnIndex = RunningList.Columns.IndexOfKey(currObj.Key);
                    //cListItem.SubItems.Add("진행중");
                    //cListItem.SubItems.Add()
                    
                    cListItem.SubItems[columnIndex].Text = currObj.Value;
                }
            }

        }

       
        public void updateListView(ListViewItem cListItem, Dictionary<string, string> updateDic)
        {

            
            //Dictionary Data 확인
            foreach (var currObj in updateDic)
            {

                //k_testcase_no키는 Listview컬럼에 포함되지 않으므로, 해당 키의 컬럼인덱스는 존재하지 않는다.
                //따라서 해당 키의 경우는 skip하고, 나머지 키는 값을 가지고 와서 ListView에 뿌려 준다.
                if (!currObj.Key.Equals(_keyList.k_testcase_no))
                {
                    System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                    //i_keyIndex = RunningList.Columns.IndexOfKey("k_status");
                    int columnIndex = RunningList.Columns.IndexOfKey(currObj.Key);
                    System.Diagnostics.Debug.WriteLine("Column Index:{0}", columnIndex);
                    cListItem.SubItems[columnIndex].Text = currObj.Value;
                }
            }

        }

        public void HeyConnect()
        {
            System.Diagnostics.Debug.WriteLine("Hey Connect");
        }

        private void txtTime1_TextChanged(object sender, EventArgs e)
        {

        }

        
       
        private void HandlGameRunning()
        {

        }

        private void HandleCartoon()
        {

        }

        private void HandleVaccine()
        {
             
        }

        private void btnBattery_Click(object sender, EventArgs e)
        {
            int batterySize = Int32.Parse(txtBattery.Text);
            _myUtility.setBatteryWH(batterySize);

            double remaining_battery = _myUtility.getBatteryWH();
            System.Diagnostics.Debug.WriteLine("Battery Exist:{0}", remaining_battery);
        }

        private void grpTestInfo_Enter(object sender, EventArgs e)
        {

        }


        //TOAN : 01/13/2019. WMP테스트 코드
        private void PlayFile(String url)
        {
            _Player = new WMPLib.WindowsMediaPlayer();
            _Player.PlayStateChange +=
                new WMPLib._WMPOCXEvents_PlayStateChangeEventHandler(Player_PlayStateChange);
            _Player.MediaError +=
                new WMPLib._WMPOCXEvents_MediaErrorEventHandler(Player_MediaError);
            _Player.URL = url;
            _Player.controls.play();
        }

       
        private void Player_PlayStateChange(int NewState)
        {
            if ((WMPLib.WMPPlayState)NewState == WMPLib.WMPPlayState.wmppsStopped)
            {
                this.Close();
            }
        }

        private void Player_MediaError(object pMediaObject)
        {
            MessageBox.Show("Cannot play media file.");
            this.Close();
        }

        private void grpRunning_Enter(object sender, EventArgs e)
        {

        }

        private void cmdCal_Click(object sender, EventArgs e)
        {
            //수학함수(Math테스트)
            double testA;
            testA = Math.Round( 12.347, 2); //소숫점 3자리에서 반올림해서 2자리 표현
            //testA = Math.Round(12.347,3);
            //testA = Math.Round(12.34, 2);
            System.Diagnostics.Debug.WriteLine("Test Value :{0} ", testA);
        }

        private void txtTime1_Click(object sender, EventArgs e)
        {

        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Click Load Button");
            System.Diagnostics.Debug.WriteLine(e.ToString());
            //this.showMessage();

            this.openFileSelectDialog();

            if (_loadURL != null)
            {
                //_xmlLoader.loadIpInfo(_loadURL);
                //_loadURL
                txtTime3.Text = _loadURL;
            }
        }

        private void openFileSelectDialog()
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "XML Files (*.avi)|*.avi|All Files (*.*)|*.*";

            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            //TOAN  :03/28/2017. 이경우는 Loop가 한번만 돌고 종료 된다.
            //Default값인 Single Selection사용.
            foreach (var path in dlg.FileNames)
            {
                //new FileInfo(path, passphraseTextBox.Text);
                System.Diagnostics.Debug.WriteLine(path.ToString());
                _loadURL = path.ToString();
            }
        }

        private void cmdStop_Click(object sender, EventArgs e)
        {
            //Terminate test forcelhy

        }

        private void cmdSet_Click(object sender, EventArgs e)
        {

        }

        private void cmdStop_Click_1(object sender, EventArgs e)
        {
            _bTerminate = true;

        }

        //TOAN : 03/20/2019. setMonitor Timer
        public void setMonitorTimer(int duration_sec)
        {
            System.Diagnostics.Debug.WriteLine("System Timer Start:{0}", this.getCurrentTime());
            _monitorTimer = new System.Timers.Timer();
            _monitorTimer.Interval = duration_sec * 1000;
            _monitorTimer.Elapsed += MonitorTimer_Elapsed;
            _monitorTimer.Start();
        }

        private void MonitorTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //System Timer등록시간 디버깅
            //System.Diagnostics.Debug.WriteLine("System Moitor Timer Start:{0}", this.getCurrentTime());
            System.Diagnostics.Debug.WriteLine("System Moitor Timer Expire:{0}", this.getCurrentTime());


            string batterylife;
            float calBattery;

            batterylife = SystemInformation.PowerStatus.BatteryLifePercent.ToString();
            calBattery = float.Parse(batterylife) * 100;

            //TOAN : 07/03/2019. DeskTop테스트의 경우 아래 조건을 무시
            if (calBattery <=_lowBattery)
            {
                if(_monitorTimer!=null)
                {
                    if(_monitorTimer.Enabled)
                    {
                        _monitorTimer.Stop();
                    }
                }

                switch (_taskRunningList)
                {
                    case TaskRunningList.TASK_WEBACTOR:
                        {
                            // _webManager.worker.RunWorkerAsync(requestMode);
                            System.Diagnostics.Debug.WriteLine("[TASK_WEBACTOR]System Moitor Timer Expire:{0}", this.getCurrentTime());
                            _webManager._testTerminate = true;
                            _webManager.worker.CancelAsync();
                            break;
                        }
                    case TaskRunningList.TASK_YOUTUBE:
                        {
                            System.Diagnostics.Debug.WriteLine("[TASK_YOUTUBE]System Moitor Timer Expire:{0}", this.getCurrentTime());
                            _youtubeManager._testTerminate = true;
                            _youtubeManager.worker.CancelAsync();
                            break;
                        }
                    case TaskRunningList.TASK_MEDIAPLAYER:
                        {
                            System.Diagnostics.Debug.WriteLine("[TASK_MEDIAPLAYER]System Moitor Timer Expire:{0}", this.getCurrentTime());
                            _videoManager._testTerminate = true;
                            _videoManager.worker.CancelAsync();
                            break;
                        }
                    case TaskRunningList.TASK_STORAGE_ACTOR:
                        {
                            System.Diagnostics.Debug.WriteLine("[TASK_STORAGE_ACTOR]System Moitor Timer Expire:{0}", this.getCurrentTime());
                            _downloadManager._testTerminate = true;
                            _downloadManager.worker.CancelAsync();
                            break;
                        }
                    case TaskRunningList.TASK_DOCUMENT:
                        {
                            System.Diagnostics.Debug.WriteLine("[TASK_DOCUMENT]System Moitor Timer Expire:{0}", this.getCurrentTime());
                            _docManager._testTerminate = true;
                            _docManager.worker.CancelAsync();
                            break;
                        }
                    default:
                        break;
                }

                //save to report
                //txtEnd.Text = this.getCurrentTime();
                //this.saveTestResult();
            }
        }

        public void finishTest()
        {
            txtEnd.Text = this.getCurrentTime();
            this.saveTestResult();
        }

       


        //Below is just test code just with timer event
        //TOAN : 06/10/2019. Battery Level이 아닌 시간으로만 체크할 때 사용하는 코드
        //TOAN : 06/30/2019. Temporary Blocking

        //TOAN : 06/30/2019. new version with Background Worker
        //개별 Task종료 Timer와 별도로 Application Test 종료조건에 실행되는 Timer
        //e.g] Lowbattery조건

        private void SystemTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //TO DO : check low battery condition
            //if(_lowbatteryCondition)
            _systemTimer.Stop();

            //TOAN : 06/30/2019. 현재 running task을 stop시킨다.
            switch (_taskRunningList)
            {
                case TaskRunningList.TASK_WEBACTOR:
                    {
                        // _webManager.worker.RunWorkerAsync(requestMode);
                        System.Diagnostics.Debug.WriteLine("System Moitor Timer Expire:{0}", this.getCurrentTime());
                        _webManager.worker.CancelAsync();
                        break;
                    }
                case TaskRunningList.TASK_YOUTUBE:
                    {
                        break;
                    }
                case TaskRunningList.TASK_MEDIAPLAYER:
                    {
                        break;
                    }
                case TaskRunningList.TASK_STORAGE_ACTOR:
                    {
                        break;
                    }
                case TaskRunningList.TASK_DOCUMENT:
                    {
                        break;
                    }
                default:
                    break;
            }
        }

       



      

        public void stopSystemTimer()
        {
            _systemTimer.Stop();
        }


        public void recordRunningList(TaskRunningList runningTask)
        {
            _taskRunningList = runningTask;
        }

        private void cmdLowBattery_Click(object sender, EventArgs e)
        {
            //int batterySize = Int32.Parse(txtBattery.Text);
            _lowBattery = Int32.Parse(txtLowBattery.Text);
        }



        ////TOAN : 03/31/2019. Code move to ReportMaker
        private void saveTestResult()
        {
            _reportMaker.reportTestResult();
        }


       
        private void cmdReport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
            {
                if(sfd.ShowDialog()==DialogResult.OK)
                {
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb = app.Workbooks.Add(XlSheetType.xlWorksheet);
                    Worksheet ws = (Worksheet)app.ActiveSheet;
                    
                    //Range 
                    app.Visible = false;
                    ws.Cells[1, 1] = "TestCase";
                    ws.Cells[1, 2] = "Status";
                    ws.Cells[1, 3] = "Remaing Battery";
                    ws.Cells[1, 4] = "Task Discharage";
                    ws.Cells[1, 5] = "Task Discharge(wh)";
                    ws.Cells[1, 6] = "Power Consumption";
                    ws.Cells[1, 7] = "Start Time";
                    ws.Cells[1, 8] = " End Time";

                    int i = 2;

                    foreach(ListViewItem item in RunningList.Items)
                    {
                        ws.Cells[i, 1] = item.SubItems[0].Text;
                        ws.Cells[i, 2] = item.SubItems[1].Text;
                        ws.Cells[i, 3] = item.SubItems[2].Text;
                        ws.Cells[i, 4] = item.SubItems[3].Text;
                        ws.Cells[i, 5] = item.SubItems[4].Text;
                        ws.Cells[i, 6] = item.SubItems[5].Text;
                        ws.Cells[i, 7] = item.SubItems[6].Text;
                        ws.Cells[i, 8] = item.SubItems[7].Text;
                        i++;
                    }
                    var rngAll = ws.UsedRange;
                    rngAll.Select();
                    rngAll.Borders.LineStyle = 1;
                    rngAll.Borders.ColorIndex = 1;
                    ws.Columns.AutoFit();

                    wb.SaveAs(sfd.FileName, XlFileFormat.xlWorkbookDefault,
                              Type.Missing,
                              Type.Missing,
                              true,
                              false,
                              XlSaveAsAccessMode.xlNoChange,
                              XlSaveConflictResolution.xlLocalSessionChanges,
                              Type.Missing,
                              Type.Missing);
                              app.Quit();

                    MessageBox.Show("Your data has been suceesfully exported.",
                                    "Message",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                    
                }

            }
        }

        private void txtTime4_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnPPT_Click(object sender, EventArgs e)
        {

        }

        private void btnMovieRank_Click(object sender, EventArgs e)
        {

        }

        private void txtTime2_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnSearchTime_Click(object sender, EventArgs e)
        {

        }

        //TOAN : 03/31/2019. Excel Report Test
        private void cmdReportTest_Click(object sender, EventArgs e)
        {
            this.saveTestResult();
        }

        //private void cmdCal_Click_1(object sender, EventArgs e)
        //{
        //    //TOAN : 04/02/2019. Task Discharge수식 검증용
        //    //20분 수행했을때 Battery가 16%씩 사용되는지 확인할 것.이건수식이 잘못된듯 하다.
        //    int batterySize;
        //    double disCharge;
        //    int batteryValue;

        //    double testVal1=15.0;
        //    double testVal2=14.0;
        //    double testResult = 0.0;

        //    testResult = testVal1 - testVal2;
        //    System.Diagnostics.Debug.WriteLine("Test Result:{0}", testResult);

        //    batterySize = int.Parse(txtInput1.Text);
        //    disCharge = double.Parse(txtInput2.Text);

        //    //Battery Utility Code테스트
        //    batteryValue = _myUtility.getBatteryLifeV1();
        //    System.Diagnostics.Debug.WriteLine("Battery Capacity:{0}", batteryValue);

        //    //소수점 2자리에서 반올림해서, 소수점 1자리로 유지
        //    disCharge = Math.Round(batterySize * (disCharge / 100), 1, MidpointRounding.AwayFromZero);
        //    System.Diagnostics.Debug.WriteLine("Battery Size:{0}, disCharge:{1}", batterySize, disCharge);

        //}

        private void btnTestModel_Click(object sender, EventArgs e)
        {
            //‪C:\autotest\report.xlsx
            //Select file with OpenFileDlg
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel files(*.xlsx;*.xlsm)|*.xlsx;*.xlsm";
            ofd.ShowDialog();

            if(ofd.FileName.Length>0)
            {
                txtTestModelResult.Text = ofd.FileName;
                _testDecision.loadExcelFile(ofd.FileName, 1);
            }else
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Select Valid FileName"));
            }

            
         }

        private void btncompareModel_Click(object sender, EventArgs e)
        {
            //‪C:\autotest\report.xlsx
            //Select file with OpenFileDlg
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "xlsx";
            ofd.Filter = "Excel files(*.xlsx;*.xlsm)|*.xlsx;*.xlsm";
            ofd.ShowDialog();

            if (ofd.FileName.Length > 0)
            {
                txtCompareModelResult.Text = ofd.FileName;
                _testDecision.loadExcelFile(ofd.FileName, 2);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Select Valid FileName"));
            }

           
        }

        private void btnDecison_Click(object sender, EventArgs e)
        {
            //TOAN : 04/07/2019. check decision. 
            _testDecision.makeDecision();
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void chkMovieRank_CheckedChanged(object sender, EventArgs e)
        {

        }
    }  //End of Form Class 

   
}
