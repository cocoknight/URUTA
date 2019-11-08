/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

    2019-04-04 : Make a test Decion Class
    2019-04-07 : compose Testcase list for src and compare
                 In case of List Collectoin is reference type. It doesn't need ref keyword for function parameter.
    2019-04-08 : Get Testcase count and get averagePower(구간별 소비전력)              
    2019-04-09 : get numeric value from string with using Regex(Regular Expression)
    2019-04-09 : compose Test information from report file(_kTestInfoList,_infoDic) 
    2019-04-09 : print Final Test Result to Excel File
    2019-04-11 : update print format for "Remaing Battery", "Task Discharge" with Percent(%)
    2019-04-18 : Bug-fix for getAveragePower (소숫점 유무에 따라 Regular Expression을 다르게 추가함)
                 "Task Discharge(wh)", "Running Time" 컬럼에 적용
    .Refernce URL
    Load Multiple Excel : https://stackoverflow.com/questions/12745783/c-sharp-read-multiple-excel-files
    C# AS캐스팅 : https://dybz.tistory.com/94
    Use FileOpenDialog : https://cmdream.tistory.com/42
    Excel Find명령 : https://www.vitoshacademy.com/c-looking-for-a-value-in-excel-with-c-visualstudio/
    C# Foreach Loop  : https://stackoverflow.com/questions/7223507/setting-range-in-for-loop
    C# parameter send : https://stackoverflow.com/questions/33471875/passing-a-list-parameter-as-ref(04-09 work)
    C# Excel Reference : https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-excel-range/204
    C# Excel Reference1 : https://www.e-iceblue.com/Tutorials/Spire.XLS/Spire.XLS-Program-Guide/Excel-Data/How-to-Align-Excel-Text-in-C.html
    Excel Cell Merge : https://stackoverflow.com/questions/532199/merging-cells-in-excel-using-c-sharp
    get_Range Error수정 : https://www.codeflair.net/2014/01/11/object-does-not-contain-a-definition-for-get_range-in-excel-c/
    Alignment 지정 : https://stackoverflow.com/questions/22535769/c-sharp-and-excellibrary-how-to-right-align-cells
--***********************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Microsoft.Office.Interop.Excel;

//TOAN : 04/08/2019. 숫자만 추출하기
using System.Text.RegularExpressions;


namespace PerformanceUsability
{
    class CTestDecision
    {
        //Limitation for Testing. 
        //Basic concept : compate test and compare model
        protected Application _app;
        protected Workbook _wb_src;
        protected Workbook _wb_dest;

        protected Workbook _wb_decision;
        protected Worksheet _ws_decision;


        protected Worksheet _ws_src;
        protected Worksheet _ws_dest;

        //Data Structure for Excel Data
        //Each row convert to Dictionary
        //Each Dictionary include to List

        //public Dictionary<string, string> _testInfoDic;
        //public List<string> _kTCColumnList;
        //protected Form1 _form1;
        protected Form1 _form1;


        protected int _startRow;
        protected int _startCol;
        protected int _currRow;
        protected int _currCol;

        //List<int> myList = new List<int>();
        //This is Default Constructor
        public List<string> _kTCColumnList;
        //public List<string> _kTCColumnList;
        public List<string> _kTestInfoList; 

        public List<Dictionary<string, string>> _tcList; //각 TestCase(Dictionry)를 저장하는 리스트 구조
        
        public List<Dictionary<string, string>> _tcListSrc;
        public List<Dictionary<string, string>> _tcListCompare;

        public Dictionary<string, string> _tcDic; //각 Testcase자료 구조(Dictionary)
        public Dictionary<string, string> _infoDicSrc; //test information을 저장할 자료구조
        public Dictionary<string, string> _infoDicDest;
        public double low_battery;

        //for test model
        public double _usagedTime;
        public double _usagedPower;
        public double _averagePower;

        //for compare model
        public double _com_usagedTime;
        public double _com_usagedPower;
        public double _com_averagePower;

        public int numofTC;
        public bool _finalDecision;

        Range _currentModel;
        Range _currentTestcase;

        public CTestDecision()
        {
            //initialize Microsoft Excel
            //_app = new Microsoft.Office.Interop.Excel.Application();

            _kTCColumnList = new List<string>();
            _kTestInfoList = new List<string>();

            _tcList = new List<Dictionary<string, string>>();
            _tcDic = new Dictionary<string, string>();

            _infoDicSrc = new Dictionary<string, string>();
            _infoDicDest = new Dictionary<string, string>();

            _tcListSrc = new List<Dictionary<string, string>>(); 
            _tcListCompare = new List<Dictionary<string, string>>();

            //TOAN : 04/07/2019. Low Battery 수식 테스트 용도
            //hard-coding이 아닌 Excel File에서 가지고 올수 있도록 변경하자.
            low_battery = 78/*83*/; 
    }

        public void connectUI(Form1 conn)
        {
            _form1 = conn;
        }

        public void loadExcelFile(string name, int type)
        {

            _app = new Microsoft.Office.Interop.Excel.Application();

            switch (type)
            {
                case 1:
                    {
                        //load test model report
                        System.Diagnostics.Debug.WriteLine("Test Model Test Report Load:{0}", name);
                        // _wb_src = _app.Workbooks.Open
                        //load excel file
                    
                        try
                        {
                            //TOAN : 04/05/2019. 접근하기 위한 다양한 방법이 있다는 것을 생각.
                            //string loadtString;
                            //Range range;
                            //TOAN : 06/11/2019. Read-Only popup을 보여주지 않기 위함.
                            _app.DisplayAlerts = false;
                            _wb_src = _app.Workbooks.Open(name);
                            _ws_src = _wb_src.Sheets[1];
                            //_ws_src = _wb_src.Worksheets.Item["Sheet1"];

                            //Cell내용 가지고 오기
                            Range range = _ws_src.Cells[4,3];
                            string loadtString = range.Value as string;
                            //string loadtString = _ws_src.Cells[4, 3]; //This is runtime error
                            //System.Diagnostics.Debug.WriteLine("System String:{0}", loadtString);

                            //TestCase record정보를 읽어와서 List Collection에 포함시켜 준다.
                            //Step1 : worksheet에서 "TestCase"가 위치한 셀을 찾는다.
                            //"TestCase"문자열 검색. TestCase문자열 다음 row부터가 Test결과이다.
                            var rngAll = _ws_src.UsedRange;
                            rngAll.Select();
                            //TOAN : 04/09/2019. Get Test Information Header
                            
                            //Get Task-Information
                            _currentModel = rngAll.Find("Model");

                            //TOAN : 04/10/2019.
                            this.composeKeyList(_ws_src, _currentModel, _kTestInfoList);
                            this.composeTestInfo(_ws_src, _currentModel, _infoDicSrc);

                            string searchResult;
                            if (_currentModel != null)
                            {
                                searchResult = "Found at \ncolumn - " + _currentModel.Column +
                                                            "\nrow - " + _currentModel.Row;
                            }
                            else
                            {
                                searchResult = "The searched string \"" +
                                        "\" is not found.";
                            }

                            //Range currentTestCase = rngAll.Find("TestCase");
                            _currentTestcase = rngAll.Find("TestCase");
                            if (_currentTestcase != null)
                            {
                                searchResult = "Found at \ncolumn - " + _currentTestcase.Column +
                                                            "\nrow - " + _currentTestcase.Row;
                            }
                            else
                            {
                                searchResult = "The searched string \""  +
                                        "\" is not found.";
                            }
                            System.Diagnostics.Debug.WriteLine("Find Result:{0}", searchResult);
                            //TOAN : 04/10/2019. key구성은 excel load할때 한번만 수행. dest을 open할때는 중복수행하면 안된다.
                            this.composeKeyList(_ws_src, _currentTestcase, _kTCColumnList);

                            this.composeTaskList(_ws_src, _currentTestcase, _tcListSrc);


                            //Debugging
                            foreach (var currTc in _tcListSrc)
                            {
                                foreach (var currObj in currTc)
                                {
                                    System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                                }
                            }

                        }
                        catch(Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("Exception :{0}", ex.ToString());
                            throw ex;
                        }
                        finally
                        {
                            _wb_src.Close();
                            _app.Quit();
                        }

                        break;
                    }

                case 2:
                    {
                        //load compare model report
                        System.Diagnostics.Debug.WriteLine("Compare Model Test Report Load:{0}", name);
                        try
                        {
                            //TOAN : 06/11/2019. Read-Only popup을 보여주지 않기 위함.
                            _app.DisplayAlerts = false;

                            _wb_dest = _app.Workbooks.Open(name);
                            _ws_dest = _wb_dest.Sheets[1];


                            var rngAll = _ws_dest.UsedRange;
                            rngAll.Select();

                            //Get Task-Information
                            _currentModel = rngAll.Find("Model");
                            string searchResult;
                            if (_currentModel != null)
                            {
                                searchResult = "Found at \ncolumn - " + _currentModel.Column +
                                                            "\nrow - " + _currentModel.Row;
                            }
                            else
                            {
                                searchResult = "The searched string \"" +
                                        "\" is not found.";
                            }
                            this.composeTestInfo(_ws_dest, _currentModel, _infoDicDest);

                            _currentTestcase = rngAll.Find("TestCase");
                            if (_currentTestcase != null)
                            {
                                searchResult = "Found at \ncolumn - " + _currentTestcase.Column +
                                                            "\nrow - " + _currentTestcase.Row;
                            }
                            else
                            {
                                searchResult = "The searched string \"" +
                                        "\" is not found.";
                            }

                            System.Diagnostics.Debug.WriteLine("Find Result:{0}", searchResult);
                            //this.composeTaskList(currentTestCase, _tcListCompare);
                            this.composeTaskList(_ws_dest, _currentTestcase, _tcListCompare);
                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {
                            _wb_dest.Close();
                            _app.Quit();
                        }
                        break;
                    }
            }
        }
        //TOAN : 04/09/2019. Thrid Version
        void composeKeyList(object ws, object keyRecord, List<string> currList)
        {
            System.Diagnostics.Debug.WriteLine("Compose Key");
            // 'H.Range(Range("C6"), Range("C6").End(xlDown)).Select'
            Worksheet sWs = ws as Worksheet;
            Range startRange = keyRecord as Range;
            Range endRange = startRange.End[XlDirection.xlToRight];
            //Range selectArea = sWs.Range[startRange, endRange].Select();
            
        
            foreach (Range ran in sWs.Range[startRange, endRange])
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Range Value:{0}", ran.Value as string));
                currList.Add(ran.Value);
            }

            System.Diagnostics.Debug.WriteLine(string.Format("List Value:{0}", currList[0])); //List의 첫번째 값 출력
        }


        //TOAN : 04/07/2019. Second Version
        void composeKeyList(object ws,object keyRecord)
        {
            System.Diagnostics.Debug.WriteLine("Compose Key");
            // 'H.Range(Range("C6"), Range("C6").End(xlDown)).Select'
            Worksheet sWs = ws as Worksheet;
            Range startRange = keyRecord as Range;
            Range endRange = startRange.End[XlDirection.xlToRight];
            var selectArea = sWs.Range[startRange, endRange].Select();

            foreach (Range ran in sWs.Range[startRange, endRange])
            {
                System.Diagnostics.Debug.WriteLine("Range Value:{0}", ran.Value as string);
                _kTCColumnList.Add(ran.Value);
            }

            System.Diagnostics.Debug.WriteLine("List Value:{0}", _kTCColumnList[0]); //List의 첫번째 값 출력
        }

        //TOAN : 04/07/2019. First Version
        void composeKeyList(object keyRecord)
        {
            System.Diagnostics.Debug.WriteLine("Compose Key");
            // 'H.Range(Range("C6"), Range("C6").End(xlDown)).Select'
            Range startRange = keyRecord as Range;
            Range endRange = startRange.End[XlDirection.xlToRight];
            var selectArea = _ws_src.Range[startRange, endRange].Select();

            foreach (Range ran in _ws_src.Range[startRange, endRange]) 
            {
                 System.Diagnostics.Debug.WriteLine("Range Value:{0}", ran.Value as string);
                _kTCColumnList.Add(ran.Value);
            }

            System.Diagnostics.Debug.WriteLine("List Value:{0}",_kTCColumnList[0]); //List의 첫번째 값 출력
        }

        void composeTaskInfo(Object targetRecord, List<Dictionary<string, string>> currList)
        {
            //필요하다면 TaskInfo정보도 TaskList와 동일하게 만들수 있다.

        }

        void composeTestInfo(Object ws, Object targetRecord, Dictionary<string,string> currDic)
        {

            Range sRange = targetRecord as Range;
            Worksheet sWs = ws as Worksheet;

            System.Diagnostics.Debug.WriteLine(string.Format("Range Row:{0},Column:{1}", sRange.Row, sRange.Column));
            //composeKeyList
            //this.composeKeyList(sWs, sRange, _kTestInfoList);

            //키에 해당하는 값을 가지고오기 위해row를 한줄 밑으로 보낸다.
            Range startRange = sWs.Cells[sRange.Row + 1, sRange.Column];
            System.Diagnostics.Debug.WriteLine("Range Row:{0},Column:{1}", startRange.Row, startRange.Column);

            Range r = startRange.End[XlDirection.xlToRight];

            int loop_index = 0;
            
            //currKeyList는 _kTestInfoList로 Hard-coding되어 있다.
            foreach (Range ran in sWs.Range[startRange, r]) //TEST OF
            {
                Object currObj = ran.Value as object;
                System.Diagnostics.Debug.WriteLine(String.Format("Real Value:{0}", currObj.ToString()));
                currDic.Add(_kTestInfoList[loop_index], currObj.ToString());
                loop_index += 1;
            }

            //Test Information print
            foreach (var currObj in currDic)
            {
                System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
            }

        }

        //Third Version(This is real version)
        //Compose Key-List확인할 것.
        void composeTaskList(Object ws,Object targetRecord, List<Dictionary<string, string>> currList)
        {
            Range sRange = targetRecord as Range;
            Worksheet sWs = ws as Worksheet;
            System.Diagnostics.Debug.WriteLine("Range Row:{0},Column:{1}", sRange.Row, sRange.Column);

            //sRange의 Cell값을 Dictionary의 키값으로 사용한다.
            //this.composeKeyList(sRange);
            //this.composeKeyList(sWs, sRange);
            //TOAN : 04/09/2019. code change
            //TOAN : 04/10/2019. composeTaskList는 Excel파일 loading할때 갱신되므로 이함수안에 있으면
            //_kTCColumnList을 두번수행하는 결과가 된다.
            //this.composeKeyList(sWs, sRange, _kTCColumnList);
            
            //실수행된 Testcase를 Dictionary List Collection에 추가 한다.
            //Range를 한줄 밑으로 옮겨서 Testcase결과를 가지고 온다.
            Range startRange = sWs.Cells[sRange.Row + 1, sRange.Column];
            System.Diagnostics.Debug.WriteLine("Range Row:{0},Column:{1}", startRange.Row, startRange.Column);

            Range r = startRange.End[XlDirection.xlDown];
            //xlDown이 적용된 제일 마지막 셀을 리턴함.
            //foreach (Range ran in _ws_src.Range["c8", r])  //TEST OK. 하지만 이경우는 range이 start가 상수("C8")이므로 적합하지 않다.

            int loop_index = 0;
            //이중 for-loop을 사용하자.
            foreach (Range ran in sWs.Range[startRange, r]) //TEST OF
            {
                //System.Diagnostics.Debug.WriteLine("Range Value:{0}", ran.Value);
                System.Diagnostics.Debug.WriteLine("Range Value:{0}", ran.Value as string);
                _tcDic = new Dictionary<string, string>();
                Range columnRange = ran.End[XlDirection.xlToRight];

                foreach (Range i_ran in sWs.Range[ran, columnRange])
                {
                    //_tcDic.Add(_kTCColumnList[loop_index], ran.Value);
                    //String.Format("LoadContent: Asset Name : {0}". theAsset))
                    // _tcDic = new Dictionary<string, string>();
                    //currKeyList는 _kTCColumnList로 HardCoding되어 있다.
                    Object currObj = i_ran.Value as object;
                    System.Diagnostics.Debug.WriteLine(String.Format("Real Value:{0}", currObj.ToString()));
                    _tcDic.Add(_kTCColumnList[loop_index], currObj.ToString());
                    loop_index += 1;

                }
                loop_index = 0;
                currList.Add(_tcDic);

            }
            //End of For-Loop

            //TOAN : 04/07/2019. List-Collection에 값을 출력하자. List Collection의 원소들은 Dictioanry이다.
            //foreach문은 Case1, Case어떤 형태도 가능하다.
            //Case 1: var형태를 사용한 루프 순회

            foreach (var currTc in currList)
            {
                foreach (var currObj in currTc)
                {
                    System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                }
            }
        }



        //Second Version
        void composeTaskList(Object targetRecord, List<Dictionary<string, string>> currList)
        {
            
                Range sRange = targetRecord as Range;
                System.Diagnostics.Debug.WriteLine("Range Row:{0},Column:{1}", sRange.Row, sRange.Column);

                //sRange의 Cell값을 Dictionary의 키값으로 사용한다.
                this.composeKeyList(sRange);

                //실수행된 Testcase를 Dictionary List Collection에 추가 한다.
                Range startRange = _ws_src.Cells[sRange.Row + 1, sRange.Column];
                System.Diagnostics.Debug.WriteLine("Range Row:{0},Column:{1}", startRange.Row, startRange.Column);

                Range r = startRange.End[XlDirection.xlDown]; 
                //xlDown이 적용된 제일 마지막 셀을 리턴함.
                //foreach (Range ran in _ws_src.Range["c8", r])  //TEST OK. 하지만 이경우는 range이 start가 상수("C8")이므로 적합하지 않다.

                int loop_index = 0;
                //이중 for-loop을 사용하자.
                foreach (Range ran in _ws_src.Range[startRange, r]) //TEST OF
                {
                    //System.Diagnostics.Debug.WriteLine("Range Value:{0}", ran.Value);
                    System.Diagnostics.Debug.WriteLine("Range Value:{0}", ran.Value as string);
                    _tcDic = new Dictionary<string, string>();
                    Range columnRange = ran.End[XlDirection.xlToRight];

                    foreach (Range i_ran in _ws_src.Range[ran, columnRange])
                    {
                        //_tcDic.Add(_kTCColumnList[loop_index], ran.Value);
                        //String.Format("LoadContent: Asset Name : {0}". theAsset))
                        // _tcDic = new Dictionary<string, string>();
                        Object currObj = i_ran.Value as object;
                        System.Diagnostics.Debug.WriteLine(String.Format("Real Value:{0}", currObj.ToString()));
                        _tcDic.Add(_kTCColumnList[loop_index], currObj.ToString());
                        loop_index += 1;

                    }
                    loop_index = 0;
                    currList.Add(_tcDic);

                }
            //End of For-Loop

            //TOAN : 04/07/2019. List-Collection에 값을 출력하자. List Collection의 원소들은 Dictioanry이다.
            //foreach문은 Case1, Case어떤 형태도 가능하다.
            //Case 1: var형태를 사용한 루프 순회

            foreach (var currTc in currList)
            {
                foreach (var currObj in currTc)
                {
                    System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                }
            }


        }


        //First Version
        void composeTaskList(Object targetRecord)
        {
            Range sRange = targetRecord as Range;
            System.Diagnostics.Debug.WriteLine("Range Row:{0},Column:{1}", sRange.Row,sRange.Column);

            //sRange의 Cell값을 Dictionary의 키값으로 사용한다.
            this.composeKeyList(sRange);

            //실수행된 Testcase를 Dictionary List Collection에 추가 한다.
            Range startRange = _ws_src.Cells[sRange.Row + 1, sRange.Column];
            System.Diagnostics.Debug.WriteLine("Range Row:{0},Column:{1}", startRange.Row, startRange.Column);

            Range r = startRange.End[XlDirection.xlDown]; //xlDown이 적용된 제일 마지막 셀을 리턴함.
                                                          //Range r = _ws_src.Range[startRange.Column& startRange.Row].End[XlDirection.xlDown]; //TEST Fail
                                                          //Range r = _ws_src.Range["c8"].End[XlDirection.xlDown]; //TEST OK
                                                          //Range r = _ws_src.Range[_ws_src.Cells[startRange.Row,startRange.Column]].End[XlDirection.xlDown];
                                                          //foreach (Range ran in _ws_src.Range["c8","c12"]) //TEST FAIL
                                                          //foreach (Range ran in _ws_src.Range["c8"].End[XlDirection.xlDown]) //이경우는 마지막  값만 가지고 온다.왜냐하면 RANGE범위가 아닌 코드에 끝값으로 포함
                                                          //foreach (Range ran in _ws_src.Range["c8", r])  //TEST OK. 하지만 이경우는 range이 start가 상수("C8")이므로 적합하지 않다.

            int loop_index = 0;
            //이중 for-loop을 사용하자.
            foreach (Range ran in _ws_src.Range[startRange, r]) //TEST OF
            {
                //System.Diagnostics.Debug.WriteLine("Range Value:{0}", ran.Value);
                System.Diagnostics.Debug.WriteLine("Range Value:{0}", ran.Value as string);
                _tcDic = new Dictionary<string, string>();
                Range columnRange = ran.End[XlDirection.xlToRight];

                foreach (Range i_ran in _ws_src.Range[ran, columnRange])
                {
                    //_tcDic.Add(_kTCColumnList[loop_index], ran.Value);
                    //String.Format("LoadContent: Asset Name : {0}". theAsset))
                    // _tcDic = new Dictionary<string, string>();
                    Object currObj = i_ran.Value as object;
                    System.Diagnostics.Debug.WriteLine(String.Format("Real Value:{0}", currObj.ToString()));
                    _tcDic.Add(_kTCColumnList[loop_index], currObj.ToString());
                    loop_index += 1;

                }
                loop_index = 0;
                _tcList.Add(_tcDic);
                
            }
            //End of For-Loop

            //TOAN : 04/07/2019. List-Collection에 값을 출력하자. List Collection의 원소들은 Dictioanry이다.
            //foreach문은 Case1, Case어떤 형태도 가능하다.
            //Case 1: var형태를 사용한 루프 순회
            foreach (var currTc in _tcList)
            {
                foreach (var currObj in currTc)
                {
                    System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                }
            }

            //Case2 : 명시적인 type-casting을 사용한 루프 순회
            //foreach (Dictionary<string, string> currTC in _tcList)
            //{
            //    foreach (KeyValuePair<string, string> kvp in currTC)
            //    {
            //        string year = kvp.Key;
            //        string month = kvp.Value;
            //        System.Diagnostics.Debug.WriteLine(string.Format("key:{0}, value:{1}", kvp.Key, kvp.Value));
            //    }
            //}


        }

        //public void makeDesision(List<Dictionary<string, string>> src, List<Dictionary<string, string>> compare)
        public void makeDecision()
        {
            //TOAN : 06/10/2019. Directory Check
            string dirName = @"C:\autotest";
            DirectoryInfo di = new DirectoryInfo(dirName);

            if (di.Exists == false)
            {
                di.Create();
            }

            //STEP1 : Compare src and compare
            //STEP2 : Low Battery기준 해당 기준보다 큰Test Case기준으로 각각의 평균 소비전력을 구한다.
            //src, compare중 low-battery도달시점까지 TC를 적게 수행한쪽을 찾는다.
            //TC를 적게 수행한쪽의 ListCount만큼 서로 동일하게 내부 TC를 수행한값을 계산한다.
            //평균소비전력 = discharge wh합계 / 수행시간의 합게 

            //Below is debugging Code
            //STEP1 : src TaskList 출력
            foreach (var currTc in _tcListSrc)
            {
                foreach (var currObj in currTc)
                {
                    System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                }
            }

            ////STEP2 : compare TaskList출력
            foreach (var currTc in _tcListCompare)
            {
                foreach (var currObj in currTc)
                {
                    System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                }
            }

            int listCountA = _tcListSrc.Count;
            int listCountB = _tcListCompare.Count;

            if(listCountA>listCountB)
            {
                //검증모델의 testcase가 low-battery시점까지 더 많이 수행되었다.
                int compareNumber = this.get_compareCount(_tcListCompare);
                System.Diagnostics.Debug.WriteLine(String.Format("Compare Number:{0}", compareNumber));
                numofTC = compareNumber - 1;
            }
            else
            {
                //비교모델의 testcase가 low-battery시점까지 더 많이 수행되었다.
                //이경우 src모델의 testcase을 기준으로 low-battery이전까지 수행항목 갯수를 체크 한다.

                int compareNumber = this.get_compareCount(_tcListSrc);
                System.Diagnostics.Debug.WriteLine(String.Format("Compare Number:{0}",compareNumber));
                numofTC = compareNumber - 1;
            }

            System.Diagnostics.Debug.WriteLine(String.Format("List Size TestMode:{0},CompareMode:{1}", listCountA,listCountB));

            _averagePower = this.getAveragePower(_tcListSrc, numofTC);
            _com_averagePower = this.getAveragePower(_tcListCompare, numofTC);

            _finalDecision = this.getfinalDecision(_averagePower,_com_averagePower);
            System.Diagnostics.Debug.WriteLine(String.Format("Final Decision:{0}", _finalDecision));


            //이제 모든 데이터가 취합되어졌으므로 결과를 출력하자.
            this.makeFinalReport(_averagePower, _com_averagePower,_finalDecision);
        }//End of makeDecision 

        void makeFinalReport(double srcAveragePower, double destAveragePower, bool decision)
        {
            //최종 판정 결과를 출력하자.
            try
            {
                _app = new Microsoft.Office.Interop.Excel.Application();
                _wb_decision = _app.Workbooks.Add(XlSheetType.xlWorksheet);
                _ws_decision = (Worksheet)_app.ActiveSheet;

                _startRow = 4;
                _startCol = 3; //C열부터 시작.

                _currRow = _startRow;
                _currCol = _startCol;

                Range startRange = _ws_decision.Cells[_currRow, _currCol];
                System.Diagnostics.Debug.WriteLine(string.Format("Start Range's Row:{0},Column:{1}", startRange.Row,startRange.Column));
             
                //Print Test Model Test Report. 아래 함수를 하나 호출했을때 검증모델의 평균소비전력 포함 출력
                this.printDecisionResult(startRange,
                                         _kTestInfoList,
                                         _kTCColumnList,
                                         _tcListSrc, 
                                         _infoDicSrc);
                //print src AveragePower
                //this.printAveragePower
                startRange = _ws_decision.Cells[_currRow, _currCol];
                int numOfCols = _kTCColumnList.Count; //testcase상세 리스트의 갯수를 가지고 온다.
                this.printAveragePower(startRange, numOfCols, srcAveragePower);

                startRange = _ws_decision.Cells[_currRow, _currCol];
                System.Diagnostics.Debug.WriteLine(string.Format("Start Range's Row:{0},Column:{1}", startRange.Row, startRange.Column));

                //Print Compare Model Test Report. 아래 함수를 하나 호출했을때 비교모델의 평균소비전력 포함 출력
                //TOAN : 04/10/2019. Temporary skip
                this.printDecisionResult(startRange,
                                         _kTestInfoList,
                                         _kTCColumnList,
                                         _tcListCompare,
                                         _infoDicDest);

                startRange = _ws_decision.Cells[_currRow, _currCol];
                this.printAveragePower(startRange, numOfCols, destAveragePower);

                //Print Final Decision
                this.printFinalDecistion(numOfCols, decision);

                //Save all worksheet fo excel file
                this.savetofile();
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Exception :{0}", ex.ToString());
                throw ex;
            }
            finally
            {
                _wb_decision.Close();
                _app.Quit();
            }

        }

        void printFinalDecistion(int area,bool finalDecision)
        {
            string displayResult;
            _ws_decision.Cells[_currRow, _currCol] = "Test Result";
            _currCol += 1;
            if (finalDecision==true)
            {
                displayResult = "PASS";
            }
            else
            {
                displayResult = "FAIL";
            }

            _ws_decision.Cells[_currRow, _currCol] = displayResult;
            int areasize = _currCol + area-/*1*/2;
            Range range = _ws_decision.get_Range((object)_ws_decision.Cells[_currRow, _currCol], (object)_ws_decision.Cells[_currRow, areasize]);
            range.Merge(true);
            range.Interior.ColorIndex = 36;

            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            _currRow += 1;
            _currCol = _startCol;
        }

        void printAveragePower(Range start,int area,double power)
        {

            //_ws_decision.Cells[_currRow, _currCol] = power.ToString() + "Wh"; 
            //Range("A2:A5").Merge
            System.Diagnostics.Debug.WriteLine(string.Format("area size:{0}", area));
            _ws_decision.Cells[_currRow, _currCol] = "Average Power Consumption";
            _currCol += 1;
            _ws_decision.Cells[_currRow, _currCol] = power.ToString() + "Wh";


            int areasize = _currCol + area - /*1*/2;
            System.Diagnostics.Debug.WriteLine(string.Format("calculation area size:{0}", areasize));
            Range range = _ws_decision.get_Range((object)_ws_decision.Cells[_currRow, _currCol], (object)_ws_decision.Cells[_currRow, areasize]);
            range.Merge(true);
            range.Interior.ColorIndex = 36;
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            //Range range = _ws_decision.Range[_ws_src.Cells[_currRow, _currCol], _ws_src.Cells[_currRow, areasize]];
            // Excel.Range range = ws.get_Range(ws.Cells[1, 1], ws.Cells[1, 2]);

            //_ws_decision.Cells[_currRow,].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            //_ws_decision.Cells[$"A{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            _currRow += 1;
            _currCol = _startCol;
        }

        void savetofile()
        {
            var rngAll = _ws_decision.UsedRange;
            rngAll.Select();
            rngAll.Borders.LineStyle = 1;
            rngAll.Borders.ColorIndex = 1;
            _ws_decision.Columns.AutoFit();

            //각 셀의 정보를 정렬하시오.
            rngAll.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            rngAll.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            var fileName = @"C:\\autotest\\testDecision.xlsx";
            if (File.Exists(fileName)) File.Delete(fileName);

            //TOAN : 04/04/2019. File save as read-write
            //아래 함수에서 fileName을 사용하면 runtime -exception발생함
            //Exception내용
            //            • 지정한 폴더가 있는지 확인하십시오. 
            //            • 파일이 들어 있는 폴더가 읽기 전용이 아닌지 확인하십시오. 
            //            • 파일 이름에<  >  ?  [  ]  : | *등의 문자가 들어 있는지 확인하십시오.
            //            • 파일이나 경로 이름은 218자를 초과할 수 없습니다.
            
            _wb_decision.SaveAs(/*fileName*/"C:\\autotest\\testDecision.xlsx", XlFileFormat.xlWorkbookDefault,
                      Type.Missing,
                      Type.Missing,
                      true,
                      false,
                      /*XlSaveAsAccessMode.xlNoChange*/XlSaveAsAccessMode.xlExclusive,
                      XlSaveConflictResolution.xlLocalSessionChanges,
                      Type.Missing,
                      Type.Missing);


            //_app.Quit();

            System.Windows.Forms.MessageBox.Show("Your data has been suceesfully exported.",
                            "Message",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
        }
        void printDecisionResult(Object startRecord,
                                 List<string> columnsInfo,
                                 List<string> tcColumnsInfo,
                                 List<Dictionary<string, string>> currList, 
                                 Dictionary<string, string> currDic)
        {
            //Step1 : Print Test Informateion(Key Information)
            //Dictionary는 순서대로 구성되어 있지 않기 때문에, List구조를 이용해서 순서대로 출력을 한다.
            foreach (string name in columnsInfo)
            {
                System.Diagnostics.Debug.WriteLine("key string:{0}", name);

                if (currDic.ContainsKey(name))
                    {
                        _ws_decision.Cells[_currRow, _currCol] = name;
                        _currCol += 1;
                    }
               
            }

            Range crange = _ws_decision.get_Range((object)_ws_decision.Cells[_currRow, _startCol], (object)_ws_decision.Cells[_currRow, /*_currCol*/_startCol+tcColumnsInfo.Count-1]);
            crange.Interior.ColorIndex = 45;


            _currRow += 1;
            _currCol = _startCol;

            foreach (string name in columnsInfo)
            {
                System.Diagnostics.Debug.WriteLine("key string:{0}", name);

                if (currDic.ContainsKey(name))
                {
                        _ws_decision.Cells[_currRow, _currCol] = currDic[name];
                        _currCol += 1;
                }
            }

            _currRow += 1;
            _currCol = _startCol;

            foreach (string name in tcColumnsInfo)
            {
                    _ws_decision.Cells[_currRow, _currCol] = name;
                    _currCol += 1;
            }

            _currRow += 1;
            _currCol = _startCol;

            //일단 전부 출력한다.
            foreach (var currTc in currList)
            {
                //curTc는 Dictioary이다.
                foreach (string name in tcColumnsInfo)
                {
                    if (currTc.ContainsKey(name))
                    {
                        //키에 해당하는 값을 출력한다.
                        //TOAN : 04/11/2019. Remaing Battery와 Task Discharge의 경우 단위(%)를 출력 한다.
                        //_ws_decision.Cells[_currRow, _currCol] = currTc[name];
                        //_currCol += 1;

                        //키에 해당하는 값을 출력한다.
                        if (name.Equals("Remaing Battery") || name.Equals("Task Discharage"))
                        {
                            double convertValue = double.Parse(currTc[name]) * 100;
                            _ws_decision.Cells[_currRow, _currCol] = convertValue.ToString() + "%";
                            _currCol += 1;
                        }
                        else
                        {
                            _ws_decision.Cells[_currRow, _currCol] = currTc[name];
                            _currCol += 1;
                        }
                    }
                }
                _currRow += 1;
                _currCol = _startCol;
            }

            //아래코드가 있으면 한줄을 더 띄우는 결과가 된다.
            //_currRow += 1;
            //_currCol = _startCol;

        }

        bool getfinalDecision(double srcAveragePower,double compareAveragePower)
        {
            bool retValue = false;
            double calValue = 0.0;
            //세번째 인자가 없으면 0.5에서 반올림이 되지 않는다.
            calValue = Math.Round(compareAveragePower * 1.1, 1, MidpointRounding.AwayFromZero);
            if(srcAveragePower>calValue)
            {
                retValue = false;
            }
            else
            {
                retValue = true;
            }
            return retValue;
        }

        int get_compareCount(List<Dictionary<string, string>> currList)
        {
            int testCount = 0;
            bool loop_exit = false;

            foreach (var currTc in _tcListSrc)
            {

                //testCount += 1;
                if (loop_exit == false)
                {
                    foreach (var currObj in currTc)
                    {
                        //TOAN : 04/08/2019. Check Remaining Battery Value
                        //txtPPTWorkingTime.Text.Equals(sPageNumber)
                        System.Diagnostics.Debug.WriteLine(string.Format("key:{0}, value:{1}", currObj.Key, currObj.Value));

                        if (currObj.Key.Equals("Remaing Battery"))
                        {
                            //int remaingBattery= currObj.Value
                            double remaingBattery = double.Parse(currObj.Value)*100;

                            if (remaingBattery==low_battery)
                            {
                                loop_exit = true;
                                break;
                            }
                        }
                        
                    }

                    testCount += 1;
                }
                else
                {
                    break;
                }

               

            }

            return testCount;
        }

        double getAveragePower(List<Dictionary<string, string>> currList,int testStep)

        {
            //call by reference전달이므로 굳이 return-type을 사용하지 않겠다.
            double calAverage=0;
            double usagedTime = 0;
            double usagedPower = 0;
            int loopCounter = 0;
            //STEP1 : testStep까지 루프 순회
            foreach (var currTc in currList)
            {
                if (loopCounter != testStep)
                {
                    foreach (var currObj in currTc)
                    {
                        System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                        if (currObj.Key.Equals("Task Discharge(wh)"))
                        {
                            
                            string strTarget = currObj.Value;
                            Regex r = new Regex(@"[0-9]+\.[0-9]+");
                            //string strTmp = Regex.Replace(strTarget, @"\D", "");
                            //double nTmp = double.Parse(strTmp);
                            //usagedPower = usagedPower + nTmp;
                            //문자열에서 실수값을 추출
                            if (strTarget.Contains("."))
                            {
                                Match m = r.Match(strTarget);
                                System.Diagnostics.Debug.WriteLine(string.Format("Match Value:{0}", m.Value));
                                double cVal = double.Parse(m.Value);
                                usagedPower = usagedPower + cVal;
                            }
                            else
                            {
                                string strTmp = Regex.Replace(strTarget, @"\D", "");
                                double nTmp = double.Parse(strTmp);
                                usagedPower = usagedPower + nTmp;
                            }

                            System.Diagnostics.Debug.WriteLine(string.Format("Usaged Power:{0}", usagedPower));
                        }

                        if (currObj.Key.Equals("Running Time"))
                        {
                            string strTarget = currObj.Value;
                            Regex r = new Regex(@"[0-9]+\.[0-9]+");
                            if (strTarget.Contains("."))
                            {
                                Match m = r.Match(strTarget);
                                double cVal = double.Parse(m.Value);
                                usagedTime = usagedTime + cVal;
                                System.Diagnostics.Debug.WriteLine(string.Format("Usaged Time:{0}", usagedTime));
                            }
                            else
                            {
                                string strTmp = Regex.Replace(strTarget, @"\D", "");
                                double nTmp = double.Parse(strTmp);
                                usagedPower = usagedPower + nTmp;
                                System.Diagnostics.Debug.WriteLine(string.Format("Usaged Time:{0}", usagedTime));
                            }
                              
                        }
                    }
                    loopCounter += 1;
                }
                else
                {
                    break;
                }

            }

            System.Diagnostics.Debug.WriteLine(string.Format("Usaged Power:{0}", usagedPower));
            System.Diagnostics.Debug.WriteLine(string.Format("Usaged Time:{0}", usagedTime));

            //STEP2 : 
            calAverage = Math.Round(usagedPower/ usagedTime, 1, MidpointRounding.AwayFromZero);
            return calAverage;
        }

        void getAveragePower(List<Dictionary<string, string>> currList, 
                               int testStep,
                               ref double usagedTime,
                               ref double usagedPower,
                               ref double averagePower)

        {
          //call by reference전달이므로 굳이 return-type을 사용하지 않겠다.

           
        }

    }   //End of Class


} //End of NameSpace
