using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//Use singletone class for KeyList Data Structure

namespace PerformanceUsability
{

    public sealed class CustomTimer
    {
        private static readonly CustomTimer instance = new CustomTimer();
        protected Form1 _uiManager;
        System.Timers.Timer _systemTimer;
         
        //생성자 초기화시 Property로 선언한 data-member값을 초기화 시켜 준다.
        private CustomTimer()
        {
           

        }

        //Data member를 숨기기 위한 방법(2) : property를 이용한 방법. Auto Implementation Propery를 사용한 방법
        public string abc = "abc";

        public static CustomTimer Instance
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

        public void setSystemTimer(int duration_sec)
        {
        

        }

        private void SystemTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
         

        }

        public void connectUI(Form1 conn)
        {
            _uiManager = conn;
            System.Diagnostics.Debug.WriteLine("connectUI(CSeleniumBase)");
            conn.HeyConnect();
        }

    }
}