using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.PointOfService;

namespace POSForDotnetOverUDMTest
{
    class POSForDotnetOverUDMTest
    {
        private static readonly string PRG = typeof(POSForDotnetOverUDMTest).Name;
        static void PrintUsage()
        {
            Console.WriteLine("Usage: {0} <DeviceType> <DeviceName>", PRG);
            Console.WriteLine("Example for PosPrinter: {0} PosPrinter PosPrinter", PRG);
            Console.WriteLine("Example for PosPrinter: {0} FiscalPrinter PosPrinterMF", PRG);
            Console.WriteLine("Return codes: ");
            Console.WriteLine("  99 = Undefined error");
            Console.WriteLine("  98 = Invalid arguments ");
            Console.WriteLine("  97 = Invalid DeviceType ");
            Console.WriteLine("  96 = Invalid DeviceName ");
            Console.WriteLine("  95 = Invalid POS for .Net configuration ");
            Console.WriteLine("  94 = The specified device is not properly configured ");
            Console.WriteLine("  93 = The specified device cannot be instantiated");
            Console.WriteLine("  1  = Open failed ");
            Console.WriteLine("  2  = Claim failed ");
            Console.WriteLine("  3  = DeviceEnable failed ");
            Console.WriteLine("  4  = Release failed ");
            Console.WriteLine("  5  = Close failed ");
            Console.WriteLine("  0  = The device is configured and ready to use ");
        }

        static int Main(string[] args)
        {
            int rc = 99; // undefined error
            PosCommon common = null;
            string deviceOpenType = String.Empty; 
            string deviceOpenName = String.Empty;
            
            try
            {
                Console.WriteLine("{0}: Test programm for Pos for .Net over UDM \n", PRG);

                if (args.Length != 2)
                {
                    PrintUsage();
                    return 98;
                }

                rc = 97;
                deviceOpenType = args[0];

                rc = 96;
                deviceOpenName = args[1];

                rc = 95;
                PosExplorer posExplorer = new PosExplorer();

                rc = 94;
                DeviceInfo device = posExplorer.GetDevice(deviceOpenType, deviceOpenName);
                if (device == null) return rc;

                rc = 93;
                common = posExplorer.CreateInstance(device) as PosCommon;
                if (common == null) return rc;

                rc = 1;
                common.Open();

                rc = 2;
                common.Claim(10000); // it is the value used in TP.Net for PosPrinter, FiscalPrinter, Display

                rc = 3;
                common.DeviceEnabled = true;

                rc = 4;
                common.Release();

                rc = 5;
                common.Close();

                rc = 0;
            }
            catch (PosException e)
            {
                Console.WriteLine("PosExcption: " + e.Message);
                if (e.InnerException != null)
                    Console.WriteLine("Inner PosExcption: " + e.InnerException.Message);
            }

            return rc;
        }
        
    }
}
