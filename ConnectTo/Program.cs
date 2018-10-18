using System;
using System.Management;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Diagnostics;
using System.Threading;


/* Reference

    https://stackoverflow.com/questions/43173970/map-network-drive-programmatically-in-c-sharp-on-windows-10?rq=1

    https://www.pinvoke.net/default.aspx/winspool.AddPrinterConnection
    https://www.pinvoke.net/default.aspx/winspool.SetDefaultPrinter

    Turn on Legacy Default Printer Mode
    https://social.technet.microsoft.com/Forums/office/en-US/e5996baa-5825-440c-940d-862a80730f8b/let-windows-manage-my-default-printer-disable-via-gpo?forum=win10itprogeneral
*/

namespace Utility
{
    public class NetworkDrive
    {
        private enum ResourceScope
        {
            RESOURCE_CONNECTED = 1,
            RESOURCE_GLOBALNET,
            RESOURCE_REMEMBERED,
            RESOURCE_RECENT,
            RESOURCE_CONTEXT
        }
        private enum ResourceType
        {
            RESOURCETYPE_ANY,
            RESOURCETYPE_DISK,
            RESOURCETYPE_PRINT,
            RESOURCETYPE_RESERVED
        }
        private enum ResourceUsage
        {
            RESOURCEUSAGE_CONNECTABLE = 0x00000001,
            RESOURCEUSAGE_CONTAINER = 0x00000002,
            RESOURCEUSAGE_NOLOCALDEVICE = 0x00000004,
            RESOURCEUSAGE_SIBLING = 0x00000008,
            RESOURCEUSAGE_ATTACHED = 0x00000010
        }
        private enum ResourceDisplayType
        {
            RESOURCEDISPLAYTYPE_GENERIC,
            RESOURCEDISPLAYTYPE_DOMAIN,
            RESOURCEDISPLAYTYPE_SERVER,
            RESOURCEDISPLAYTYPE_SHARE,
            RESOURCEDISPLAYTYPE_FILE,
            RESOURCEDISPLAYTYPE_GROUP,
            RESOURCEDISPLAYTYPE_NETWORK,
            RESOURCEDISPLAYTYPE_ROOT,
            RESOURCEDISPLAYTYPE_SHAREADMIN,
            RESOURCEDISPLAYTYPE_DIRECTORY,
            RESOURCEDISPLAYTYPE_TREE,
            RESOURCEDISPLAYTYPE_NDSCONTAINER
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct NETRESOURCE
        {
            public ResourceScope oResourceScope;
            public ResourceType oResourceType;
            public ResourceDisplayType oDisplayType;
            public ResourceUsage oResourceUsage;
            public string sLocalName;
            public string sRemoteName;
            public string sComments;
            public string sProvider;
        }
        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2
            (ref NETRESOURCE oNetworkResource, string sPassword,
            string sUserName, int iFlags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2
            (string sLocalName, uint iFlags, int iForce);

        public static int MapNetworkDrive(string sDriveLetter, string sNetworkPath)
        {
            //Checks if the last character is \ as this causes error on mapping a drive.
            if (sNetworkPath.Substring(sNetworkPath.Length - 1, 1) == @"\")
            {
                sNetworkPath = sNetworkPath.Substring(0, sNetworkPath.Length - 1);
            }

            NETRESOURCE oNetworkResource = new NETRESOURCE()
            {
                oResourceType = ResourceType.RESOURCETYPE_DISK,
                sLocalName = sDriveLetter + ":",
                sRemoteName = sNetworkPath
            };

            //If Drive is already mapped disconnect the current 
            //mapping before adding the new mapping
            string currentNetworkPath = GetCurrentMapping(sDriveLetter);

            if (currentNetworkPath == "")
            {
                // Not connected do nothing
            }
            else if (currentNetworkPath == sNetworkPath)
            {
                // Already connected
                return 0;
            }
            else {
                // Connected to something else. Disconnect first
                DisconnectNetworkDrive(sDriveLetter, true);
            }

            return WNetAddConnection2(ref oNetworkResource, null, null, 0);
        }

        public static int DisconnectNetworkDrive(string sDriveLetter, bool bForceDisconnect)
        {
            if (bForceDisconnect)
            {
                return WNetCancelConnection2(sDriveLetter + ":", 0, 1);
            }
            else
            {
                return WNetCancelConnection2(sDriveLetter + ":", 0, 0);
            }
        }


        public static string GetCurrentMapping(string sDriveLetter)
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                "select * from Win32_MappedLogicalDisk where caption = '" + sDriveLetter + ":'");
            foreach (ManagementObject drive in searcher.Get())
            {
                return drive["ProviderName"].ToString();
            }
            return "";
        }
    }

    public class Printer
    {
        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetDefaultPrinter(string Name);

        [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AddPrinterConnection(String pszBuffer);
    }

}


namespace ConnectTo
{
    class Program
    {
        static void PrintUsage()
        {
            Console.WriteLine("Usage: ConnectTo (-printer|-defaultprinter|-share letter) resource");
            Environment.Exit(1);
        }

        static void LogInformation(string message)
        {
            /* using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = "Application";
                eventLog.WriteEntry(message, EventLogEntryType.Information);
            } */
            EventLog.WriteEntry(".NET Runtime", message, EventLogEntryType.Information, 1000);
        }
        static void LogError(string message)
        {
            /* using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = "Application";
                eventLog.WriteEntry(message, EventLogEntryType.Error);
            } */
            EventLog.WriteEntry(".NET Runtime", message, EventLogEntryType.Error, 1000);
        }

        static void ConnectToPrinter(string printer, bool defaultPrinter)
        {
            int tryNo = 1;
            while (tryNo < 4) {
                bool success = ConnectToPrinterAux(printer, defaultPrinter, tryNo);
                if (success)
                {
                    Environment.Exit(0);
                }
                tryNo++;
                Thread.Sleep(5000);
            }
            Environment.Exit(1);
        }

        static bool ConnectToPrinterAux(string printer, bool defaultPrinter, int tryNo)
        {
            int error;
            LogInformation("connectTo " + (defaultPrinter ? "-defaultprinter" : "-printer") + " " + printer + " try #" + tryNo);
            bool success = Utility.Printer.AddPrinterConnection(printer);
            
            if (! success)
            {
                error = Marshal.GetLastWin32Error();
                LogError("connectTo AddPrinterConnection exit code = " + error);
                return false;
            }
            if (defaultPrinter)
            {
                // Turn on legacy default printer mode
                Registry.SetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows", "LegacyDefaultPrinterMode", 1);
                success = Utility.Printer.SetDefaultPrinter(printer);
                if (!success)
                {
                    error = Marshal.GetLastWin32Error();
                    LogError("connectTo SetDefaultPrinter exit code = " + error);
                    return false;
                }
            }
            return true;
        }

        static void ConnectToShare(string driveLetter, string share)
        {
            LogInformation("connectTo -share " + driveLetter + " " + share);
            // Console.WriteLine("Connecting " + driveLetter + " to " + share);
            if (driveLetter.Length > 1)
            {
                driveLetter = driveLetter.Substring(0, 1);
            }
            driveLetter = driveLetter.ToUpper();
            if (driveLetter.CompareTo("D") == -1 || driveLetter.CompareTo("Z") == 1)
            {
                LogError("connectTo letter " + driveLetter + ": not allowed");
                Environment.Exit(1);
            }

            int errorCode = Utility.NetworkDrive.MapNetworkDrive(driveLetter, share);
            if (errorCode != 0)
            {
                LogError("connectTo MapNetworkDrive exit code = " + errorCode);
            }
            Environment.Exit(errorCode);
        }

        static void Main(string[] args)
        {
            if (args.Length == 2 && args[0].Equals("-printer"))
            {
                ConnectToPrinter(args[1], false);
            }
            else if (args.Length == 2 && args[0].Equals("-defaultprinter"))
            {
                ConnectToPrinter(args[1], true);
            }
            else if (args.Length == 3 && args[0].Equals("-share"))
            {
                ConnectToShare(args[1], args[2]);
            }
            else
            {
                PrintUsage();
            }
        }
    }
}
