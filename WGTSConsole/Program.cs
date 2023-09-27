using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
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
        public static int serial_fail_count = 0;
        public static string logPath = "C:\\WGTSLog\\log.log";
        public static Stopwatch stopwatch = new Stopwatch();

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
            if (!Directory.Exists("C:\\WGTSLog"))
            {
                Directory.CreateDirectory("C:\\WGTSLog");
            }

            File.AppendAllText(logPath, "\n=== WGTS Sync Start, Before Sync ===");
            File.AppendAllText(logPath, '\n' + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffffff"));

            stopwatch.Start();

            serial.DataReceived += new SerialDataReceivedEventHandler(serial_DataReceived);
            serial.ErrorReceived += new SerialErrorReceivedEventHandler(serial_ErrorReceived);

            while (!serial_gps_received)
            {
                Thread.Sleep(3000);

                if (!serial_gps_received)
                {
                    serial.Close();
                    serial_OpenPort();

                    serial_port_number += 1;
                }

                serial_fail_count += 1;

                // Timeout, tyring w32tm
                if (serial_fail_count >= 5)
                {
                    Console.WriteLine("Timeout, trying Windows time service - 재시도 시간 초과, Windows 시간 동기를 시도합니다.");
                    File.AppendAllText(logPath, '\n' + "Timeout, trying Windows time service - 재시도 시간 초과, Windows 시간 동기를 시도합니다.");

                    ProcessStartInfo pri = new ProcessStartInfo();
                    Process pro = new Process();

                    pri.FileName = "cmd.exe";
                    pri.UseShellExecute = false;
                    pri.Arguments = "/c net start w32time & w32tm -config -update & w32tm -resync -rediscover & w32tm -resync -nowait";
                    pri.RedirectStandardOutput = true;

                    pro.StartInfo = pri;
                    pro.Start();
                    string output = pro.StandardOutput.ReadToEnd();

                    Console.WriteLine(output);
                    File.AppendAllText(logPath, '\n' + output);

                    pro.WaitForExit();

                    Environment.Exit(0);
                }
            }
        }

        private static void serial_OpenPort()
        {
            string[] ports = SerialPort.GetPortNames();

            if (ports.Length == 0)
            {
                Console.WriteLine("GPS port not found - GPS가 연결된 시리얼포트를 찾지 못했습니다. (재시도 중)");
                File.AppendAllText(logPath, '\n' + "GPS port not found - GPS가 연결된 시리얼포트를 찾지 못했습니다. (재시도 중)");

                Thread.Sleep(3000);
                return;
            }

            if (ports.Length <= serial_port_number)
            {
                serial_port_number = 0;
            }

            Console.WriteLine("Discovered COM ports : " + ports.Length + ", Trying COM port #" + serial_port_number.ToString() + "...");
            Console.WriteLine("발견 포트 개수 : " + ports.Length + ", " + serial_port_number.ToString() + "번 포트 시도 중...");
            File.AppendAllText(logPath, '\n' + "발견 포트 개수 : " + ports.Length + ", " + serial_port_number.ToString() + "번 포트 시도 중...");

            if (ports.Length != 0)
            {
                serial.PortName = ports[serial_port_number];
                serial.BaudRate = 9600;
                serial.StopBits = StopBits.One;
                serial.Parity = Parity.None;
                serial.DataBits = 8;

                if (serial.IsOpen)
                {
                    Console.WriteLine("COM ports already in use - 이미 사용 중인 COM 포트입니다.");
                    File.AppendAllText(logPath, '\n' + "COM ports already in use - 이미 사용 중인 COM 포트입니다.");
                    return;
                }
                else
                {
                    serial.Open();
                }
            }
            else
            {
                Console.WriteLine("No available COM ports - 연결된 시리얼포트를 찾지 못했습니다.");
                File.AppendAllText(logPath, '\n' + "No available COM ports - 연결된 시리얼포트를 찾지 못했습니다.");
                return;
            }
        }

        private static void serial_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            File.AppendAllText(logPath, '\n' + "시리얼 에러 탐지 -> " + e.ToString());

            serial_port_number += 1;

            Console.WriteLine(serial_port_number);
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

                        stopwatch.Stop();
                        TimeSpan ts = stopwatch.Elapsed;

                        Console.WriteLine("Time synchronization success - 시간 동기화 성공");
                        File.AppendAllText(logPath, '\n' + "Time synchronization success - 시간 동기화 성공");

                        File.AppendAllText(logPath, "\n=== WGTS Sync Elapsed Time===");
                        File.AppendAllText(logPath, '\n' + String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                            ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10));
                        File.AppendAllText(logPath, "\n=== WGTS Sync Finished, After Sync ===");
                        File.AppendAllText(logPath, '\n' + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffffff"));

                        serial.Close();
                        Thread.Sleep(3000);
                        Environment.Exit(0);
                        return;
                    }
                }
            }
            catch (Exception)
            {
                g_sRecvData = string.Empty;
                Console.WriteLine("Unable to synchronize time - 시간 동기화가 정상적으로 이루어지지 않았습니다.");
                File.AppendAllText(logPath, '\n' + "Unable to synchronize time - 시간 동기화가 정상적으로 이루어지지 않았습니다. (재시도 중)");
                return;
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
