/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

    2019-01-02 : Make a SingleTone Class
    2019-03-31 : Add key string for CReportMaker
    2019-04-04 : Add key for Total Time 
    
--***********************************************************************************************************/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//Use singletone class for KeyList Data Structure

namespace PerformanceUsability
{

    public sealed class KeyList
    {
        private static readonly KeyList instance = new KeyList();

        //생성자 초기화시 Property로 선언한 data-member값을 초기화 시켜 준다.
        private KeyList()
        {
            System.Diagnostics.Debug.WriteLine("Call By Static Function");
            k_testcase_no = "k_testcase_no";
            k_testcase = "k_testcase";
            k_status = "k_status";
            k_remaining_battery = "k_remaining_battery";
            k_discharge = "k_discharge";
            k_discharge_wh = "k_discharge_wh";
            k_power_consumption_wh = "k_power_consumption";
            k_start_time = "k_start_time";
            k_end_time = "k_end_time";
            //TOAN : 04/04/2019.
            k_running_time = "k_running_time";

            //TOAN : 03/31/2019.
            k_test_category = "Test Information";
            k_test_model = "Model";
            k_test_battery_wh = "Battery(Wh)";
            k_test_start_time = "Start Time";
            k_test_end_time = "End Time";
            k_test_start_battery = "Start Battery(%)";
            k_test_low_battery = "Low Battery(%)";

           
           
        }

    

        //Data member를 숨기기 위한 일반적인 방법(1): private data member를 선언하고 set/get멤버 펑션을 별도 구현하는것
        //const is constant.
        //private  const string k_testcase = "k_testcase";
        //private  const string k_status = "k_status";
        //private  const string k_remaining_battery = "k_remaining_battery";
        //private  const string k_discharge = "k_discharge";
        //private  const string K_discharge_wh = "k_discharge_wh";
        //private  const string K_power_consumption = "k_power_consumption";


        //Data member를 숨기기 위한 방법(2) : property를 이용한 방법. Auto Implementation Propery를 사용한 방법
        public string abc = "abc";

        public static KeyList Instance
        {
            get 
            {
                return instance;
            }
        }

        public string getTC()
        {
            return abc;
        }

        //C# Property. Normal Propery Declaration
        //public string K_TestCase
        //{
        //    get { return k_testcase;}
        //set { k_testcase = value; }
        //}

        //C# Auto Implementation Property. Just get property is valid.
        //별도의 Data member선언없이 property형태로 선언과 동시에 처리.별도의 private멤버를 선언할 필요 없이
        //은닉성을 보장해 준다.

        public string k_testcase_no
        {
            get;
            //set;
        }

        public string k_testcase
        {
            get;
            //set;
        }

        public string k_status
        {
            get;
            //set;
        }

        public string k_remaining_battery
        {
            get;
            //set;
        }

        public string k_discharge
        {
            get;
            //set;
        }

        public string k_discharge_wh
        {
            get;
            //set;
        }

      
        public string k_power_consumption_wh
        {
            get;
            //set;
        }


        public string k_start_time
        {
            get;
            //set;
        }

        public string k_end_time
        {
            get;
            //set;
        }

        public string k_running_time
        {
            get;
            //set;
        }

        //TOAN : 03/31/2019. Add Key for CReportMaker
        public string k_test_category
        {
            get;
        }

        public string k_test_model
        {
            get;
        }

        public string k_test_battery_wh
        {
            get;
        }

        public string k_test_start_time
        {
            get;
        }

        public string k_test_end_time
        {
            get;
        }

        public string k_test_start_battery
        {
            get;
        }

        public string k_test_low_battery
        {
            get;
        }

    }
}
