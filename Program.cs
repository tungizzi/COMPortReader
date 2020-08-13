using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.IO.Ports;
using System.Threading;

namespace COMPortReader
{
    public class SerialPortReader
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

        public class DataPackage
        {
            public static int ReceiveCount = 0;
            public static DataTable dataTable = new DataTable();

            public static void ReceiptCount(string source)
            {
                if (!dataTable.Columns.Contains("Receipt count"))
                {
                    dataTable.Columns.Add("Receipt count", typeof(int));
                }
                if (source.Contains("Receive Finished"))
                {
                    ReceiveCount++;
                }

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

            public static string GetData(string source, string keyword1, string keyword2, string col)
            {
                if (!dataTable.Columns.Contains(col))
                {
                    dataTable.Columns.Add(col, typeof(string));
                }
                if (source.Contains(keyword1))
                {
                    string dt = getBetween(source, keyword1, keyword2);
                    //Int32.TryParse(dt, out int dtVal);
                    //DataRow row = dataTable.NewRow();
                    //row[col] = dtVal;
                    //dataTable.Rows.InsertAt(row, dataTable.Columns.IndexOf(col));
                    return dt;
                }else
                {
                    return null;
                }
            }

            public static void ExtractData(string source)
            {
                string[] data = new string[4];
                data[0] = GetData(source, "RSSI:", "dBm", "RSSI (dBm)");
                data[1]= GetData(source, "SNR:", "dB", "SNR (dB)");
                data[2]= GetData(source, "Payload_size:", "bytes", "Payload Size (bytes)");
                data[3] = GetData(source, "Payload_data:", ":End_payload_data", "Payload Data");
                DataRow row = dataTable.NewRow();
                for(int i = 0; i < 4; i++)
                {
                    row[i] = data[i];
                }
                dataTable.Rows.Add(row);
            }

                public static void ExportExcel()
            {
                DataRow countReceipt = dataTable.NewRow();
                countReceipt["Receipt count"] = ReceiveCount;
                dataTable.Rows.Add(countReceipt);

                IXLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(dataTable, "result");
                wb.SaveAs(SaveDirectory + ExcelName);
            }
        }

        public static void Read()
        {
            while (_continue)
            {
                try
                {
                    string message = _serialPort.ReadLine();
                    Console.WriteLine(message);
                    DataPackage.ExtractData(message);
                    /*Dp.GetData(message, "Payload number", ";", "Payload number");
                    Dp.GetData(message, "RSSI", "dBm", "RSSI (dBm)");
                    Dp.GetData(message, "SNR", "dB", "SNR (dB)");
                    Dp.GetData(message, "Payload size", "bytes", "Payload Size (bytes)");
                    Dp.GetData(message, "Payload data", ";", "Payload Data");*/
                    DataPackage.ReceiptCount(message);
                }
                catch (TimeoutException) { }
            }
            DataPackage.ExportExcel();
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