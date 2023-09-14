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
        public static bool serial_gps_received = false;
        public static int serial_port_number = 0;
        public static int counter = 0;
        public static int counterForceCloseLimit = 10;

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

            int counterForceClose = 0;

            while (true)
            {
                if (counter >= 5)
                {
                    serial.Close();
                    Environment.Exit(0);
                }

                Thread.Sleep(1000);

                counterForceClose++;

                if (counterForceClose >= counterForceCloseLimit)
                {
                    serial.Close();
                    Console.WriteLine("Time expired - 시간 초과로 종료합니다.");
                    Environment.Exit(-2);
                }

                if (!serial_gps_received)
                {
                    serial.Close();
                    serial_OpenPort();
                    serial_port_number += 1;
                }
            }
        }

        private static void serial_OpenPort()
        {
            string[] ports = SerialPort.GetPortNames();

            if (ports.Length < serial_port_number + 1)
            {
                Console.WriteLine("GPS port not found - GPS가 연결된 시리얼포트를 찾지 못했습니다.");
                Thread.Sleep(3000);
                Environment.Exit(-2);
            }

            Console.WriteLine("Discovered COM ports : " + ports.Length + ", Trying COM port #" + serial_port_number.ToString() + "...");
            Console.WriteLine("발견 포트 개수 : " + ports.Length + ", " + serial_port_number.ToString() + "번 포트 시도 중...");

            if (ports.Length != 0)
            {
                serial.PortName = ports[serial_port_number];
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
                Console.WriteLine("No available COM ports - 연결된 시리얼포트를 찾지 못했습니다.");
                Environment.Exit(-2);
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
                    serial_gps_received = true;
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

                        Console.WriteLine("Time synchronization success - 시간 동기화 성공");

                        serial.Close();
                        Thread.Sleep(3000);
                        Environment.Exit(0);
                        return;
                    }

                    counter++;
                }
            }
            catch (Exception)
            {
                g_sRecvData = string.Empty;
                Console.WriteLine("Unable to synchronize time - 시간 동기화가 정상적으로 이루어지지 않았습니다.");
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
