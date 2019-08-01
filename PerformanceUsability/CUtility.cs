/*********************************************************************************************************-- 
    
    Copyright (c) 2019, YongMin Kim. All rights reserved. 
    This file is licenced under a Creative Commons license: 
    http://creativecommons.org/licenses/by/2.5/ 

    2019-01-10   : Make a SingleTone Class
    2019-04-03 : add new verion of getBatteryLife with interger return type
    2019-08-01 : change battery variable int -> double
--***********************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//TOAN : 10/14/2018. Process Start, Like as Media Player
using System.Diagnostics;

//TOAN : 12/12/2018. Get Power Information with Battery
using System.Management;
using System.Windows.Forms;

namespace PerformanceUsability
{
    class CUtility
    {
        //singletone data-member
        private static readonly CUtility instance = new CUtility();

        //protected string _batteryLife;
        protected string _currentTime;
        //protected int _batterySize;
        protected double _batterySize;
        private CUtility()
        {

        }

        //singletone static property
        public static CUtility Instance
        {
            get
            {
                return instance;
            }
        }

        public double getBatteryLife()
        {

            double calBattery;
            string batterylife;

            batterylife = SystemInformation.PowerStatus.BatteryLifePercent.ToString();
            //calBattery = Int32.Parse(batterylife);
            calBattery = double.Parse(batterylife) * 100;
            //MessageBox.Show(calBattery.ToString());
            //txtCurrentBattery.Text = calBattery.ToString();
            //_batteryLife = calBattery.ToString() + "%";

            return calBattery;
        }

        public int getBatteryLifeV1()
        {

            double calBattery;
            string batterylife;
            int con_batterylife;

            batterylife = SystemInformation.PowerStatus.BatteryLifePercent.ToString();
            //calBattery = Int32.Parse(batterylife);
            calBattery = double.Parse(batterylife) * 100;
            con_batterylife = Convert.ToInt32(calBattery);
            return con_batterylife;
        }



        public string getCurrentTime()
        {
            //Get Current Time
            string currTime;
            //startTime = string.Format("{0:hh:mm:ss tt}", DateTime.Now);
            currTime = string.Format("{0:hh:mm tt}", DateTime.Now);
            _currentTime = currTime;
            return _currentTime;
        }

        //public void setBatteryWH(int batterySize)
        public void setBatteryWH(double batterySize)
        {
            _batterySize = batterySize;
        }

        //public int getBatteryWH()
        public double getBatteryWH()
        {
            return _batterySize;
        }
        //string batterylife;
        ////int calBattery;
        //float calBattery;

        //batterylife = SystemInformation.PowerStatus.BatteryLifePercent.ToString();
        //    //calBattery = Int32.Parse(batterylife);
        //    calBattery = float.Parse(batterylife)*100;
        //    MessageBox.Show(calBattery.ToString());

        //    txtCurrentBattery.Text = calBattery.ToString();
    }
}
