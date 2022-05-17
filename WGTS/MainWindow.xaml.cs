using System;
using System.Diagnostics;
using System.Globalization;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;

namespace WGTS
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SerialPort serial = new SerialPort();
        private string g_sRecvData = String.Empty;
        public static bool isClosing = false;

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

        public MainWindow()
        {
            InitializeComponent();

            Loaded += new RoutedEventHandler(InitSerialPort);
        }

        private void InitSerialPort(object sender, EventArgs e)
        {
            serial.DataReceived += new SerialDataReceivedEventHandler(serial_DataReceived);

            string[] ports = SerialPort.GetPortNames();

            foreach (string port in ports)
            {
                cbComPort.Items.Add(port);
            }
        }

        private void OpenComPort(object sender, RoutedEventArgs e)
        {
            try
            {
                serial.Open();
            }
            catch
            {
                cbComPort.SelectedItem = "";
            }
        }

        private void serial_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                while (true)
                {
                    if (isClosing)
                    {
                        serial.Close();
                        break;
                    }

                    g_sRecvData = String.Empty;
                    try
                    {
                        g_sRecvData = serial.ReadLine();
                    }
                    catch
                    {
                        break;
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

                            setWindowsTime(dt);
                        }

                        Dispatcher.Invoke(
                            System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate
                            {
                                tbRecvData.Text = g_sRecvData;
                                tbConvertedTime.Text = utcfull;
                            }));
                    }
                }
            }
            catch (TimeoutException)
            {
                g_sRecvData = string.Empty;
            }
        }

        private void setWindowsTime(DateTime dt)
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

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            serial.DataReceived += new SerialDataReceivedEventHandler(serial_DataReceived);

            string[] ports = SerialPort.GetPortNames();

            if (ports.Length < 1)
            {
                cbComPort.Items.Clear();
            }
            else
            {
                cbComPort.Items.Clear();
                foreach (string port in ports)
                {
                    cbComPort.Items.Add(port);
                }
            }
        }

        private void cbComPort_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (serial.IsOpen)
            {
                serial.Close();
            }

            if (cbComPort.SelectedItem != null)
            {
                serial.PortName = cbComPort.SelectedItem.ToString();
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
                    OpenComPort(sender, e);
                }
            }
        }
        private void Window_Closing(object sender, EventArgs e)
        {
            isClosing = true;
        }
    }
}
