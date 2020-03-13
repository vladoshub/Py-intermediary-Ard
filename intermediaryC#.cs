using System;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.IO.Ports;
using System.Threading;


namespace ConsoleApp1
{


    class Start
    {

        public static void Main(string[] args)
        {
            Program program = new Program();
            program.Open();
            while (true)
            {
                program.Transfer();
                Thread.Sleep(200);
            }
        }

    }



    class Program
    {
        string port = "";
        SerialPort currentPort;
        int coord;
        string vid="1A86", pid="7523";
        public void Open()
        {
       
            string port = GetCOM(vid, pid);
            if (!port.Equals(""))
            {
                currentPort = new SerialPort(port, 115200, Parity.None, 8, StopBits.One);
                Thread.Sleep(500);
                try
                {
                    currentPort.Open();
                }
                catch {

                }
            

            }
        }

        private void ComToStd()
        {
          
                string S = currentPort.ReadLine();
                Console.WriteLine(S);

        }

        private void Array()
        {

            string S = currentPort.ReadLine();
            Console.WriteLine(S);
            for (int i = 0; i < int.Parse(S); i++)
            {
                ComToStd();
                ComToStd();
            }

        }


        public void Transfer()
        {
          
                string Ops = Console.ReadLine();
                switch (Ops)
                {
                    case "N":
                    currentPort.Write(Ops);
                    ComToStd();
                    break;
                    case "W":
                    currentPort.Write(Ops);
                    break;
                    case "M":
                    currentPort.Write(Ops);
                    Array();
                    break;
                    case "C":
                    currentPort.Write(Ops);
                    break;
                    case "S":
                    currentPort.Write(Ops);
                    currentPort.Write(Console.ReadLine());
                    currentPort.Write(Console.ReadLine());
                    break;
                    case "E":
                    currentPort.Write(Ops);
                    Environment.Exit(0);
                    break;
                    case "T":
                    currentPort.Write(Ops);
                    ComToStd();
                    break;
                }
    

        }


       public void setVP(string vid,string pid)
        {
            this.vid = vid;
            this.pid = pid;
        }


        public string GetCOM(string vid, string pid)
        {
            List<string> names = ComPortNames(vid, pid);
            if (names.Count > 0)
            {
                foreach (String s in SerialPort.GetPortNames())
                {
                    if (names.Contains(s))
                    {
                        return s;
                    }
                }
                return "";
            }
            else
                return "";
        }


        List<string> ComPortNames(String VID, String PID)
        {
            String pattern = String.Format("^VID_{0}.PID_{1}", VID, PID);
            Regex _rx = new Regex(pattern, RegexOptions.IgnoreCase);
            List<string> comports = new List<string>();
            RegistryKey rk1 = Registry.LocalMachine;
            RegistryKey rk2 = rk1.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum");
            foreach (String s3 in rk2.GetSubKeyNames())
            {
                RegistryKey rk3 = rk2.OpenSubKey(s3);
                foreach (String s in rk3.GetSubKeyNames())
                {
                    if (_rx.Match(s).Success)
                    {
                        RegistryKey rk4 = rk3.OpenSubKey(s);
                        foreach (String s2 in rk4.GetSubKeyNames())
                        {
                            RegistryKey rk5 = rk4.OpenSubKey(s2);
                            RegistryKey rk6 = rk5.OpenSubKey("Device Parameters");
                            comports.Add((string)rk6.GetValue("PortName"));
                        }
                    }
                }
            }
            return comports;
        }

    }

}
