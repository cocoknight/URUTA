/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

    2019-03-30 : Make a test report as wanted Format
    -First report format as Excel
    
    2019-03-31 : use C# Collection
    Reference URL - https://mrw0119.tistory.com/18
    Reference URL - http://www.csharp-examples.net/foreach/

    2019-04-04 : Add running time
    2019-04-04 : Save file with RW attribute
    2019-04-09 : change report content. remove "Test Information" string
    2019-04-09 : Exception Handling for Report Make
    2019-06-24 : display a Message Box on top of all forms.
    2022-01-11 : Add TestReport Result Item
    -Average Power Consumption
    -Total Running Time
--***********************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Threading;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
//using System.Windows.Forms;

namespace PerformanceUsability
{
    class CReportMaker
    {
        public Dictionary<string, string> _testInfoDic;
        public List<string> _kTCColumnList;
        protected Form1 _form1;

        //get keylist instance
        protected KeyList _keyList;

        protected string s_test_category;
        protected string s_test_model;
        protected string s_test_battery_wh;
        protected string s_test_start_time;
        protected string s_test_end_time;
        protected string s_test_start_battery;
        protected string s_test_low_battery;


        //C#의 Data-Member는 선언과 동시에 초기화가 필요 없지만, 지역변수의 경우는 선언과함께 초기화기 필요하다.
        //그렇지 않으면 run-time error가 발생 한다.
        //Microsoft.Office.Interop.Excel.Application

        protected Application _app ;
        protected Workbook _wb;
        protected Worksheet _ws;

        protected int _startRow;
        protected int _startCol;
        protected int _currRow;
        protected int _currCol;

        //TOAN : 01/11/2022. Total Average Power, Total Running Time추가
        public double _usagedTime;
        public double _usagedPower;
        public double _averagePower;

        public double _usagedRunningTime;
        public double _totalRunningTime;


        //This is Default Constructor
        public CReportMaker()
        {
            _kTCColumnList = new List<string>(); 
            _testInfoDic = new Dictionary<string, string>();
            _keyList = KeyList.Instance;

            //initialize Microsoft Excel
            //TOAN : 04/09/2019. Code-change
             //_app = new Microsoft.Office.Interop.Excel.Application();
             //_wb = _app.Workbooks.Add(XlSheetType.xlWorksheet);
             //_ws = (Worksheet)_app.ActiveSheet;

            _startRow = 4;
            _startCol = 3; //C열부터 시작.

            _currRow = _startRow;
            _currCol = _startCol;
        }

       
          
        public void reportTestInformation()
         {
            //step1 : compose List with keys
            //TOAN : 04/09/2019. remove k_test_category from report
            _kTCColumnList.Add(_keyList.k_test_category);
            _kTCColumnList.Add(_keyList.k_test_model);
            _kTCColumnList.Add(_keyList.k_test_battery_wh);
            _kTCColumnList.Add(_keyList.k_test_start_time);
            _kTCColumnList.Add(_keyList.k_test_end_time);
            _kTCColumnList.Add(_keyList.k_test_start_battery);
            _kTCColumnList.Add(_keyList.k_test_low_battery);

            //step1 : compose Dictionary
            _testInfoDic.Clear();

            s_test_category = _form1.grpTestInfo.Text;
            s_test_model = _form1.txtModel.Text;
            s_test_battery_wh = _form1.txtBattery.Text;
            s_test_start_time = _form1.txtStart.Text;
            s_test_end_time =_form1.txtEnd.Text;
            s_test_start_battery =_form1.txtCurrentBattery.Text;
            s_test_low_battery=_form1.txtLowBattery.Text;

            _testInfoDic.Add(_keyList.k_test_category, s_test_category);
            _testInfoDic.Add(_keyList.k_test_model, s_test_model);
            _testInfoDic.Add(_keyList.k_test_battery_wh, s_test_battery_wh);
            _testInfoDic.Add(_keyList.k_test_start_time, s_test_start_time);
            _testInfoDic.Add(_keyList.k_test_end_time, s_test_end_time);
            _testInfoDic.Add(_keyList.k_test_start_battery, s_test_start_battery);
            _testInfoDic.Add(_keyList.k_test_low_battery, s_test_low_battery);
       
            //step2 : print dictionary to excel
            //2-1 : Print Test Category
            //TOAN : 04/09/2019. remove Test Information from report
            //_ws.Cells[_startRow, _startCol] = _testInfoDic[_keyList.k_test_category];

            _currRow += 1;

            //2-2 : print key column
            foreach (string name in _kTCColumnList)
            {
                //Console.WriteLine(name);
                System.Diagnostics.Debug.WriteLine("key string:{0}", name);
   
                if (!name.Equals(_keyList.k_test_category))
                {
                    if (_testInfoDic.ContainsKey(name))
                    {
                        _ws.Cells[_currRow, _currCol] = name/*_testInfoDic[name].ToString()*/;
                        _currCol += 1;
                    }
                }
            }

            _currRow += 1;
            _currCol = _startCol;

            //2-3 : Print key value
            foreach (string name in _kTCColumnList)
            {
                //System.Diagnostics.Debug.WriteLine("key:{0}, value:{1}", currObj.Key, currObj.Value);
                if (!name.Equals(_keyList.k_test_category))
                {
                    if (_testInfoDic.ContainsKey(name))
                    {
                        _ws.Cells[_currRow, _currCol] = _testInfoDic[name].ToString();
                        _currCol += 1;
                    }
                }
            }

            //var rngAll = _ws.UsedRange;
            //rngAll.Select();
            //rngAll.Borders.LineStyle = 1;
            //rngAll.Borders.ColorIndex = 1;
            //_ws.Columns.AutoFit();

            //var fileName = @"C:\\autotest\\report.xlsx";
            //if (File.Exists(fileName)) File.Delete(fileName);

            //_wb.SaveAs("C:\\autotest\\report.xlsx", XlFileFormat.xlWorkbookDefault,
            //          Type.Missing,
            //          Type.Missing,
            //          true,
            //          false,
            //          XlSaveAsAccessMode.xlNoChange,
            //          XlSaveConflictResolution.xlLocalSessionChanges,
            //          Type.Missing,
            //          Type.Missing);


            //_app.Quit();

            //System.Windows.Forms.MessageBox.Show("Your data has been suceesfully exported.",
            //                "Message",
            //                System.Windows.Forms.MessageBoxButtons.OK,
            //                System.Windows.Forms.MessageBoxIcon.Information);

            //TOAN : 01/12/2021. 사용이 끝났으면 다시 _currCol값을 바꾼다.
            _currCol = _startCol;
        }


        //TOAN : 01/11/2022. 시험결과 항목 업데이트
        //검증모델의 시험 결과에 "Average Power Consumption"과 "Total Running Time" Field추가.

        //public void reportTask
        public void reportTaskResult()
         {

            //TOAN : 01/09/2018. 컬럼추가시 Key를 이용한 방법으로 변경
            //RunningList.Columns.Add(_keyList.k_testcase, "TestCase", /*300*/250);
            //RunningList.Columns.Add(_keyList.k_status, "Status", 80);
            //RunningList.Columns.Add(_keyList.k_remaining_battery, "Remaing Battery", 90);
            //RunningList.Columns.Add(_keyList.k_discharge, "Task Discharge", 90);
            //RunningList.Columns.Add(_keyList.k_discharge_wh, "Task Discharge(wh)", 110);
            //RunningList.Columns.Add(_keyList.k_power_consumption_wh, "Power Consumption", 110);
            //RunningList.Columns.Add(_keyList.k_start_time, "Start Time", 75);
            //RunningList.Columns.Add(_keyList.k_end_time, "End Time", 80);
            //RunningList.Columns.IndexOfKey(currObj.Key);

            //TOAN : 04/03/2019. Test Information다음에 Test결과를 보여 준다.
            //Test결과 Header정보 출력
               _currRow += 1;
               _ws.Cells[_currRow, _startCol] = "TestCase";
               _ws.Cells[_currRow, _startCol+1] = "Status";
               _ws.Cells[_currRow, _startCol+2] = "Remaing Battery";
               _ws.Cells[_currRow, _startCol+3] = "Task Discharage";
               _ws.Cells[_currRow, _startCol+4] = "Task Discharge(wh)";
               _ws.Cells[_currRow, _startCol+5] = "Power Consumption";
               _ws.Cells[_currRow, _startCol+6] = "Start Time";
               _ws.Cells[_currRow, _startCol+7] = "End Time";
               _ws.Cells[_currRow, _startCol + 8] = "Running Time";
            //Test결과 정보 출력
            _currRow += 1;

            //TOAN : 01/12/2022. save number of columns
            int col_num = 9;

            foreach (System.Windows.Forms.ListViewItem item in _form1.RunningList.Items)
            {
                 col_num = item.SubItems.Count;
                _ws.Cells[_currRow, _startCol] = item.SubItems[0].Text;
                _ws.Cells[_currRow, _startCol+1] = item.SubItems[1].Text;
                _ws.Cells[_currRow, _startCol+2] = item.SubItems[2].Text;
                _ws.Cells[_currRow, _startCol+3] = item.SubItems[3].Text;
                _ws.Cells[_currRow, _startCol+4] = item.SubItems[4].Text;

                //TOAN : 01/11/2022. accumulate item.SubItems[4].Text value for Average Power Consumption
                //column name : Task Discharge(wh)
                //구분자 오류가 없게 하기 위해 일괄적으로 대문자로 변환
                string strPowerWH = item.SubItems[4].Text.ToUpper();
                strPowerWH = Regex.Replace(strPowerWH, "WH", "");
                double dPowerWH = Double.Parse(strPowerWH);
                _usagedPower += dPowerWH;

                _ws.Cells[_currRow, _startCol+5] = item.SubItems[5].Text;
                _ws.Cells[_currRow, _startCol+6] = item.SubItems[6].Text;
                _ws.Cells[_currRow, _startCol+7] = item.SubItems[7].Text;
                _ws.Cells[_currRow, _startCol+8] = item.SubItems[8].Text;

                //TOAN : 01/11/2022. accumulate
                //column name : Running Time
                //구분자 오류가 없게하기 위해 일괄적으로 대문자로 변환
                string runningTime = item.SubItems[8].Text.ToUpper();
                runningTime = Regex.Replace(runningTime, "HR", "");
                double dRunningTime = Double.Parse(runningTime);
                _usagedTime += dRunningTime;

                _currRow += 1;
            }

            //TOAN : 01/11/2022. Add new rows and print "Average Power Consumption" , "Total Running Time"
            System.Diagnostics.Debug.WriteLine("[Report-Maker]Column Nmber:{0}", col_num);

            _averagePower = Math.Round(_usagedPower / _usagedTime, 1, MidpointRounding.AwayFromZero);
            System.Diagnostics.Debug.WriteLine("[Report-Maker]Total usaged power[WH]:{0}", _usagedPower);
            System.Diagnostics.Debug.WriteLine("[Report-Maker]Total running time[HR]:{0}", _usagedTime);
            System.Diagnostics.Debug.WriteLine("[Report-Maker]Average Power Consumption[WH]:{0}", _averagePower);



            //TOAN : 01/11/2022. 엑셀파일에 출력 하기
            string columnValue = "";
            Range startRange = _ws.Cells[_currRow, _startCol];
            System.Diagnostics.Debug.WriteLine(string.Format("Start Range's Row:{0},Column:{1}", startRange.Row, startRange.Column));

            //startRange = _ws_decision.Cells[_currRow, _currCol];
            //int numOfCols = _kTCColumnList.Count; //testcase상세 리스트의 갯수를 가지고 온다.
            columnValue = "Average Power Consumption";
            this.printAccumulateValue(columnValue,startRange, col_num, _averagePower,"Wh");

            //_currRow += 1;
            //_currCol = _startCol;

            startRange = _ws.Cells[_currRow, _startCol];
            columnValue = "Total Running Time";
            this.printAccumulateValue(columnValue, startRange,col_num, _usagedTime,"hr");
        }

        ////TOAN  01/11/2022. display average power
        void printAccumulateValue(string colName,Range start, int area, double power,string unit)
        {

            //_ws_decision.Cells[_currRow, _currCol] = power.ToString() + "Wh"; 
            //Range("A2:A5").Merge
            System.Diagnostics.Debug.WriteLine(string.Format("area size:{0}", area));
            _ws.Cells[_currRow, _currCol] = colName;
            _currCol += 1;
            _ws.Cells[_currRow, _currCol] = power.ToString() + unit;


            int areasize = _currCol + area - /*1*/2;
            System.Diagnostics.Debug.WriteLine(string.Format("calculation area size:{0}", areasize));
            Range range = _ws.get_Range((object)_ws.Cells[_currRow, _currCol], (object)_ws.Cells[_currRow, areasize]);
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





        public void savetofile()
        {
            var rngAll = _ws.UsedRange;
            rngAll.Select();
            rngAll.Borders.LineStyle = 1;
            rngAll.Borders.ColorIndex = 1;
            _ws.Columns.AutoFit();

            var fileName = @"C:\\autotest\\report.xlsx";
            if (File.Exists(fileName)) File.Delete(fileName);

            //TOAN : 04/04/2019. File save as read-write
            _wb.SaveAs("C:\\autotest\\report.xlsx", XlFileFormat.xlWorkbookDefault,
                      Type.Missing,
                      Type.Missing,
                      true,
                      false,
                      /*XlSaveAsAccessMode.xlNoChange*/XlSaveAsAccessMode.xlExclusive,
                      XlSaveConflictResolution.xlLocalSessionChanges,
                      Type.Missing,
                      Type.Missing);


            //TOAN : 04/11/2019. app close루틴은 finally 코드로 옮김
            //_app.Quit();
            //System.Windows.Forms.MessageBox.Show()


            //System.Windows.Forms.MessageBox.Show("Your data has been suceesfully exported.",
            //                "Message",
            //                System.Windows.Forms.MessageBoxButtons.OK,
            //                System.Windows.Forms.MessageBoxIcon.Information);

            //TOAN : 06/25/2019. 
            //https://stackoverflow.com/questions/11910448/displaying-a-messagebox-on-top-of-all-forms-setting-location-and-or-color
            //https://076923.github.io/posts/C-25/

            //System.Windows.Forms.MessageBox.Show(_form1,"Your data has been suceesfully exported.",
            //                "Message",
            //                System.Windows.Forms.MessageBoxButtons.OK,
            //                System.Windows.Forms.MessageBoxIcon.Information);


            //System.Windows.Forms.MessageBox.Show("Configuration file was corrupted.\n\nDo you want to reset it to default and lose all configurations?",
            //"Message", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question,
            // System.Windows.Forms.MessageBoxDefaultButton.Button2, System.Windows.Forms.MessageBoxOptions.ServiceNotification);

            System.Windows.Forms.MessageBox.Show("Your data has been suceesfully exported.",
            "Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information,
             System.Windows.Forms.MessageBoxDefaultButton.Button2, System.Windows.Forms.MessageBoxOptions.ServiceNotification);
        }
        public void reportTestResult()
        {
            //TOAN : 04/09/2019. Add Exception Handling
            try
            {
                _app = new Microsoft.Office.Interop.Excel.Application();
                _wb = _app.Workbooks.Add(XlSheetType.xlWorksheet);
                _ws = (Worksheet)_app.ActiveSheet;

                this.reportTestInformation();
                this.reportTaskResult();
                this.savetofile();
            }catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Full Stacktrace: {0}", ex.ToString()));
            }
            finally
            {
                _wb.Close();
                _app.Quit();
            }
        }

        public void connectUI(Form1 conn)
        {
            _form1 = conn;
        }

    }
}
