using System;
using System.Data;
using System.IO;
using System.IO.Ports;
using System.Threading;
using ClosedXML.Excel;

namespace COMPortReader
{
    public class PortChat
    {
        private static bool _continue;
        private static SerialPort _serialPort;
        static string ExcelName;
        static string SaveDirectory;

        public static void Main()
        {
            string message;
            
            StringComparer stringComparer = StringComparer.OrdinalIgnoreCase;

            Thread readThread = new Thread(Read);

            // Create a new SerialPort object with default settings.
            _serialPort = new SerialPort();

            // Allow the user to set the appropriate properties.
            _serialPort.PortName = SetPortName(_serialPort.PortName);

            // Set the read/write timeouts
            _serialPort.ReadTimeout = 500;
            _serialPort.WriteTimeout = 500;

            // Set excel name
            ExcelName = SetExcelName("test.xlsx");
            SaveDirectory = SetDirectory(Directory.GetCurrentDirectory());

            _serialPort.Open();
            _continue = true;
            readThread.Start();

            Console.WriteLine("Type QUIT to exit");

            while (_continue)
            {
                message = Console.ReadLine();

                if (stringComparer.Equals("quit", message))
                {
                    _continue = false;
                }
            }

            readThread.Join();
            _serialPort.Close();
            Console.WriteLine(Directory.GetCurrentDirectory());
        }

        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            const int kNotFound = -1;

            var startIdx = strSource.IndexOf(strStart);
            if (startIdx != kNotFound)
            {
                startIdx += strStart.Length;
                var endIdx = strSource.IndexOf(strEnd, startIdx);
                if (endIdx > startIdx)
                {
                    return strSource.Substring(startIdx, endIdx - startIdx);
                }
            }
            return String.Empty;
        }

        static int ReceiveCount = 0;

        public static void Read()
        {
            DataTable mytable = new DataTable();
            mytable.Columns.Add("SNR (unit)", typeof(int));
            while (_continue)
            {
                try
                {
                    string message = _serialPort.ReadLine();
                    Console.WriteLine(message);


                    if (message.Contains("SNR "))
                    {
                        string SNRstring = getBetween(message, "SNR ", " Payload");
                        int SNRval;
                        Int32.TryParse(SNRstring, out SNRval);
                        mytable.Rows.Add(SNRval);
                    }

                    if (message.Contains("Receive Finished"))
                    {
                        ReceiveCount++;
                    }

                }
                catch (TimeoutException) { }
            }
            mytable.Columns.Add("Receipt count", typeof(int));
            DataRow countReceipt = mytable.NewRow();
            countReceipt["Receipt count"] = ReceiveCount;
            mytable.Rows.Add(countReceipt);
            var wb = new XLWorkbook();
            wb.Worksheets.Add(mytable, "result");
            wb.SaveAs(SaveDirectory + ExcelName);
        }

        public static string SetPortName(string defaultPortName)
        {
            string portName;

            Console.WriteLine("Available Ports:");
            foreach (string s in SerialPort.GetPortNames())
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("COM port({0}): ", defaultPortName);
            portName = Console.ReadLine();

            if (portName == "")
            {
                portName = defaultPortName;
            }
            return portName;
        }

        public static string SetExcelName(string defaultExcelName)
        {
            string ExcelName;

            Console.Write("Excelname(DefaultName: {0}): ", defaultExcelName);
            ExcelName = Console.ReadLine() + ".xlsx";

            if (ExcelName == "")
            {
                ExcelName = defaultExcelName;
            }
            return ExcelName;
        }

        public static string SetDirectory(string defaultDirectory)
        {
            string Dirlink;

            Console.Write("Set directory(Default: {0}): ", defaultDirectory);
            Dirlink = Console.ReadLine() + "/";

            if (Dirlink == "")
            {
                Dirlink = defaultDirectory;
            }
            return Dirlink;
        }
    }
}