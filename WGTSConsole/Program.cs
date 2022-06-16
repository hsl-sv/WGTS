using System;
using System.Globalization;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;

namespace WGTSConsole
{
    public class Program
    {
        private static SerialPort serial = new SerialPort();
        private static string g_sRecvData = String.Empty;
        public static int counter = 0;
        public static int counterForceCloseLimit = 5;

        [StructLayout(LayoutKind.Sequential)]
        public struct SYSTEMTIME
        {
            public short wYear;
            public short wMonth;
            public short wDayOfWeek;
            public short wDay;
            public short wHour;
            public short wMinute;
            public short wSecond;
            public short wMilliseconds;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetSystemTime(ref SYSTEMTIME st);

        public static void Main(string[] args)
        {
            serial.DataReceived += new SerialDataReceivedEventHandler(serial_DataReceived);

            string[] ports = SerialPort.GetPortNames();

            if (ports.Length != 0)
            {
                serial.PortName = ports[0];
                serial.BaudRate = 9600;
                serial.StopBits = StopBits.One;
                serial.Parity = Parity.None;
                serial.DataBits = 8;

                if (serial.IsOpen)
                {
                    serial.Close();
                }
                else
                {
                    serial.Open();
                }
            }
            else
            {
                Console.WriteLine("연결된 시리얼포트를 찾지 못했습니다.");
                Environment.Exit(-2);
            }

            int counterForceClose = 0;

            while (true)
            {
                if (counter >= 3)
                {
                    serial.Close();
                    Environment.Exit(0);
                }

                Thread.Sleep(1500);

                counterForceClose++;

                if (counterForceClose >= counterForceCloseLimit)
                {
                    Console.WriteLine("시간 초과로 종료합니다.");
                    Environment.Exit(-2);
                }
            }
        }

        private static void serial_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                g_sRecvData = String.Empty;

                try
                {
                    g_sRecvData = serial.ReadLine();
                }
                catch
                {
                    return;
                }

                var recvData = g_sRecvData.Split(",");
                string recvString = String.Empty;
                string utcfull = string.Empty;

                //iRcvData += System.Text.Encoding.ASCII.GetByteCount(g_sRecvData);

                if ((g_sRecvData != string.Empty) && (recvData[0] == "$GPRMC"))
                {
                    string utctime = recvData[1];    // UTC Time hhmmss.sss
                    string status = recvData[2]; // Valid/Invalid
                    string lat = recvData[3];    // ddmm.mmmm
                    string ind_ns = recvData[4]; // N/S indicator
                    string lon = recvData[5];    // dddmm.mmmm
                    string ind_ew = recvData[6]; // E/W indicator

                    string utcdate = recvData[9]; // UTC Date DDMMYY

                    DateTime dt;
                    recvString = g_sRecvData;

                    if (utcdate != String.Empty && utctime != String.Empty)
                    {
                        string dtraw = utcdate + " " + utctime;
                        DateTime.TryParseExact(dtraw, "ddMMyy HHmmss.ff", CultureInfo.CurrentCulture,
                            DateTimeStyles.None, out dt);
                        utcfull = "UTC: " + dt.ToString();
                        Console.WriteLine(utcfull);

                        setWindowsTime(dt);
                    }

                    counter++;
                }
            }
            catch (Exception)
            {
                g_sRecvData = string.Empty;
                Console.WriteLine("시간 동기화가 정상적으로 이루어지지 않았습니다.");
                Environment.Exit(-1);
            }
        }

        private static void setWindowsTime(DateTime dt)
        {

            SYSTEMTIME st = new SYSTEMTIME();

            st.wYear = (short)dt.Year;
            st.wMonth = (short)dt.Month;
            st.wDay = (short)dt.Day;
            st.wHour = (short)dt.Hour;
            st.wMinute = (short)dt.Minute;
            st.wSecond = (short)dt.Second;
            st.wMilliseconds = (short)dt.Millisecond;

            // UTC automatically adjust to timezone of PC (on Windows)
            SetSystemTime(ref st);
        }
    }
}
