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
        static int count = 0;

        private static bool _continue;
        private static SerialPort _serialPort;

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
            Console.WriteLine("{0},{1}",count,Directory.GetCurrentDirectory());
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

        public static void Read()
        {
            DataTable mytable = new DataTable();
            mytable.Columns.Add("Interrupt", typeof(string));
            while (_continue)
            {
                try
                {
                    string message = _serialPort.ReadLine();
                    Console.WriteLine(message);


                    if (message.Contains("Interrupt"))
                    {
                        string SNRval = getBetween(message, "Interrupt", "happens");
                        mytable.Rows.Add(SNRval);
                    }

                    if (message.Contains("Receive Finished"))
                    {
                        count++;
                    }

                }
                catch (TimeoutException) { }
            }
            mytable.Columns.Add("Receipt count", typeof(int));
            DataRow countReceipt = mytable.NewRow();
            countReceipt["Receipt count"] = count;
            mytable.Rows.Add(countReceipt);
            var wb = new XLWorkbook();
            wb.Worksheets.Add(mytable, "test");
            string directory = Directory.GetCurrentDirectory();
            wb.SaveAs(directory + "test.xlsx");
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
    }
}