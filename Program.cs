using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.IO.Ports;
using System.Threading;

namespace COMPortReader
{
    public class PortChat
    {
        private static bool _continue;
        private static SerialPort _serialPort;
        private static string ExcelName;
        private static string SaveDirectory;

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

            // Set excel name and save location
            SaveDirectory = SetDirectory(Directory.GetDirectoryRoot(Directory.GetCurrentDirectory()));
            ExcelName = SetExcelName("test.xlsx");

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

        private static int ReceiveCount = 0;
        private static DataTable dataTable = new DataTable();

        public static void InitDataTable()
        {
            dataTable.Columns.Add("SNR (unit)", typeof(int));
            dataTable.Columns.Add("Receipt count", typeof(int));
        }

        public static void ExtractData(string source)
        {
            if (source.Contains("SNR "))
            {
                string SNRstring = getBetween(source, "SNR ", " Payload");
                Int32.TryParse(SNRstring, out int SNRval);
                dataTable.Rows.Add(SNRval);
            }

            if (source.Contains("Receive Finished"))
            {
                ReceiveCount++;
            }
        }

        public static void ExportExcel(string location, DataTable table)
        {
            DataRow countReceipt = dataTable.NewRow();
            countReceipt["Receipt count"] = ReceiveCount;
            dataTable.Rows.Add(countReceipt);

            IXLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(table, "result");
            wb.SaveAs(location);
        }

        public static void Read()
        {
            InitDataTable();
            while (_continue)
            {
                try
                {
                    string message = _serialPort.ReadLine();
                    Console.WriteLine(message);
                    ExtractData(message);
                }
                catch (TimeoutException) { }
            }
            ExportExcel(SaveDirectory + ExcelName, dataTable);
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