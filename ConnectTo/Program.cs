using System;
using System.Management;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Diagnostics;
using System.Threading;
using System.Text;


/* Reference

    https://stackoverflow.com/questions/43173970/map-network-drive-programmatically-in-c-sharp-on-windows-10?rq=1

    https://www.pinvoke.net/default.aspx/winspool.AddPrinterConnection
    https://www.pinvoke.net/default.aspx/winspool.SetDefaultPrinter

    Turn on Legacy Default Printer Mode
    https://social.technet.microsoft.com/Forums/office/en-US/e5996baa-5825-440c-940d-862a80730f8b/let-windows-manage-my-default-printer-disable-via-gpo?forum=win10itprogeneral
*/

namespace Utility
{
    public enum ErrorCodes
    {
        NO_ERROR = 0x0,
        ERROR_ACCESS_DENIED = 0x5,
        ERROR_BAD_DEV_TYPE = 0x42,
        ERROR_BAD_NET_NAME = 0x43,
        ERROR_ALREADY_ASSIGNED = 0x55,
        ERROR_INVALID_PASSWORD = 0x00000056,
        ERROR_BUSY = 0x000000aa,
        ERROR_BAD_DEVICE = 0x4B0,
        ERROR_CONNECTION_UNAVAIL = 0x4B1,
        ERROR_BAD_PROFILE = 0x4b6,
        ERROR_NOT_CONNECTED = 0x8CA,
        ERROR_OPEN_FILES = 0x961,
        ERROR_DEVICE_ALREADY_REMEMBERED = 0x000004b2,
        ERROR_NO_NET_OR_BAD_PATH = 0x000004b3,
        ERROR_CANNOT_OPEN_PROFILE = 0x000004b5,
        ERROR_EXTENDED_ERROR = 0x000004b8,
        ERROR_NO_NETWORK = 0x000004c6,
        ERROR_CANCELLED = 0x000004c7
    }
   
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
        [DllImport("mpr.dll")]
        public static extern int WNetGetConnection(
            string localName,
            StringBuilder remoteName,
            ref int length);

        public static string GetErrorMessage(int ErrorCode)
        {
            string Message = "";
            switch (ErrorCode)
            {
                case (int)ErrorCodes.ERROR_ACCESS_DENIED:
                    Message = "Access is denied.";
                    break;
                case (int)ErrorCodes.ERROR_BAD_DEV_TYPE:
                    Message = "The network resource type is not correct.";
                    break;
                case (int)ErrorCodes.ERROR_BAD_NET_NAME:
                    Message = "The network name cannot be found.";
                    break;
                case (int)ErrorCodes.ERROR_ALREADY_ASSIGNED:
                    Message = "The local device name is already in use.";
                    break;
                case (int)ErrorCodes.ERROR_INVALID_PASSWORD:
                    Message = "The specified network password is not correct.";
                    break;
                case (int)ErrorCodes.ERROR_BUSY:
                    Message = "The requested resource is in use.";
                    break;
                case (int)ErrorCodes.ERROR_BAD_DEVICE:
                    Message = "The specified device name is invalid.";
                    break;
                case (int)ErrorCodes.ERROR_CONNECTION_UNAVAIL:
                    Message = "The device is not currently connected but it is a remembered connection.";
                    break;
                case (int)ErrorCodes.ERROR_BAD_PROFILE:
                    Message = "The network connection profile is corrupted.";
                    break;
                case (int)ErrorCodes.ERROR_NOT_CONNECTED:
                    Message = "This network connection does not exist.";
                    break;
                case (int)ErrorCodes.ERROR_OPEN_FILES:
                    Message = "This network connection has files open or requests pending.";
                    break;
                case (int)ErrorCodes.ERROR_DEVICE_ALREADY_REMEMBERED:
                    Message = "The local device name has a remembered connection to another network resource.";
                    break;
                case (int)ErrorCodes.ERROR_NO_NET_OR_BAD_PATH:
                    Message = "The network path was either typed incorrectly, does not exist, or the network provider is not currently available. Please try retyping the path or contact your network administrator.";
                    break;
                case (int)ErrorCodes.ERROR_CANNOT_OPEN_PROFILE:
                    Message = "Unable to open the network connection profile.";
                    break;
                case (int)ErrorCodes.ERROR_EXTENDED_ERROR:
                    Message = "An extended error has occurred.";
                    /*
                    WNetGetLastErrorA(
                        LPDWORD lpError,
                        LPSTR   lpErrorBuf,
                        DWORD   nErrorBufSize,
                        LPSTR   lpNameBuf,
                        DWORD   nNameBufSize
                        );
                    */
                    break;
                case (int)ErrorCodes.ERROR_NO_NETWORK:
                    Message = "The network is not present or not started.";
                    break;
                case (int)ErrorCodes.ERROR_CANCELLED:
                    Message = "The operation was canceled by the user.";
                    break;
                case (int)ErrorCodes.NO_ERROR:
                    Message = "No error.";
                    break;
                default:
                    Message = string.Format("{0}", ErrorCode);
                    break;
            }
            return Message;
        }

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
            string currentNetworkPath = GetCurrentWNetMapping(sDriveLetter);

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

        public static string GetCurrentWNetMapping(string sDriveLetter)
        {
            int length = 250;
            StringBuilder currentUNC = new StringBuilder(length);
            int result = WNetGetConnection(sDriveLetter + ":", currentUNC, ref length);
            Console.WriteLine(GetErrorMessage(result));
            if (result != (int)ErrorCodes.NO_ERROR)
            {
                Console.WriteLine(currentUNC.ToString());
                return "";
            }
            else
            {
                Console.WriteLine(currentUNC.ToString());
                return currentUNC.ToString();
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
        public static int ConnectToShare (string sDriveLetter, string sShare)
        {
            int length = 300;
            int result = 0;
            int timeToWait = 5000;
            int numberOfTries = 10;
            int i = 0;
            StringBuilder currentShare = new StringBuilder(length);
            while (!currentShare.ToString().Equals(sShare) && i < numberOfTries)
            {
                // Get current mapping
                result = WNetGetConnection((sDriveLetter + ":"), currentShare, ref length);
                
                if (result == (int)ErrorCodes.NO_ERROR && currentShare.ToString().Equals(sShare))
                {
                    Logger.LogInformation(string.Format("{1}: is connected to {1}",sDriveLetter,sShare));
                    return result;
                }
                /*
                 * Cancelling a disconnected persistent connection results in ERROR_CONNECTION_UNAVAIL
                 * the connection needs to be restored before it can be cancelled
                 */
                else if (result == (int)ErrorCodes.ERROR_CONNECTION_UNAVAIL)
                {
                    NETRESOURCE oNetworkResource = new NETRESOURCE()
                    {
                        oResourceType = ResourceType.RESOURCETYPE_DISK,
                        sLocalName = sDriveLetter + ":",
                        sRemoteName = currentShare.ToString()
                    };
                    result = WNetAddConnection2(ref oNetworkResource, null, null, 1);
                    Logger.LogInformation(string.Format("Add old connection operation gave {0} as result. Message: {1}", result, GetErrorMessage(result)));
                }
                // Not connected, connect it!
                else if (result == (int)ErrorCodes.ERROR_NOT_CONNECTED)
                {
                    NETRESOURCE oNetworkResource = new NETRESOURCE()
                    {
                        oResourceType = ResourceType.RESOURCETYPE_DISK,
                        sLocalName = sDriveLetter + ":",
                        sRemoteName = sShare
                    };
                    result = WNetAddConnection2(ref oNetworkResource, null, null, 1);
                    Logger.LogInformation(string.Format("Add operation gave {0} as result. Message: {1}", result, GetErrorMessage(result)));
                    //Console.WriteLine(GetErrorMessage(tes2));
                }
                /* 
                    NO_ERROR means it retrieved the connection successfully. 
                    If it does not match the path it should cancel the existing connection.
                */
                //else if (result == (int)ErrorCodes.ERROR_CONNECTION_UNAVAIL || result == (int)ErrorCodes.NO_ERROR && !currentShare.ToString().Equals(sShare))
                else if (result == (int)ErrorCodes.NO_ERROR && !currentShare.ToString().Equals(sShare))
                        {
                    //Logger.LogInformation(string.Format("Local path ({0}) is connected to '{1}' but should be connected to '{2}'", sDriveLetter + ":", currentShare.ToString(), sShare));
                    result = WNetCancelConnection2(sDriveLetter + ":", 1, 1);
                    Logger.LogInformation(string.Format("Cancel operation gave {0} as result. Message: {1}", result, GetErrorMessage(result)));
                }

                Logger.LogInformation(string.Format("Last operation gave {0} as result. Message: {1}",result,GetErrorMessage(result)));
                Thread.Sleep(timeToWait);
                i++;
            }
            //Console.WriteLine("Tried {0} times out of {1}, waited for {2} seconds", i, numberOfTries, i * timeToWait / 1000);
            return result;
        }

    }

    public class Logger
    {
        public static void LogInformation(string message)
        {
            /* using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = "Application";
                eventLog.WriteEntry(message, EventLogEntryType.Information);
            } */
            EventLog.WriteEntry(".NET Runtime", message, EventLogEntryType.Information, 1000);
        }
        public static void LogError(string message)
        {
            /* using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = "Application";
                eventLog.WriteEntry(message, EventLogEntryType.Error);
            } */
            EventLog.WriteEntry(".NET Runtime", message, EventLogEntryType.Error, 1000);
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
            Console.WriteLine("Usage: ConnectTo (-printer|-defaultprinter|-share letter|-share letter name) resource");
            Environment.Exit(1);
        }
        /*
        static void LogInformation(string message)
        {
            // using (EventLog eventLog = new EventLog("Application"))
            //{
            //    eventLog.Source = "Application";
            //    eventLog.WriteEntry(message, EventLogEntryType.Information);
            //}
            EventLog.WriteEntry(".NET Runtime", message, EventLogEntryType.Information, 1000);
        }
        static void LogError(string message)
        {
            //using (EventLog eventLog = new EventLog("Application"))
            //{
            //    eventLog.Source = "Application";
            //    eventLog.WriteEntry(message, EventLogEntryType.Error);
            //} 
            EventLog.WriteEntry(".NET Runtime", message, EventLogEntryType.Error, 1000);
        }
       */
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
            Utility.Logger.LogInformation(String.Format("connectTo {0} {1} try #{2}", (defaultPrinter ? "-defaultprinter" : "-printer"), printer, tryNo));
            bool success = Utility.Printer.AddPrinterConnection(printer);
            
            if (! success)
            {
                error = Marshal.GetLastWin32Error();
                Utility.Logger.LogError(String.Format("connectTo AddPrinterConnection exit code = {0}", error));
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
                    Utility.Logger.LogError(String.Format("connectTo SetDefaultPrinter exit code = {0}", error));
                    return false;
                }
            }
            return true;
        }

        /*
        static void ConnectToShare(string driveLetter, string share)
        {
            LogInformation(String.Format("connectTo -share {0} {1}", driveLetter, share));
            // Console.WriteLine("Connecting " + driveLetter + " to " + share);
            if (driveLetter.Length > 1)
            {
                driveLetter = driveLetter.Substring(0, 1);
            }
            driveLetter = driveLetter.ToUpper();
            if (driveLetter.CompareTo("D") == -1 || driveLetter.CompareTo("Z") == 1)
            {
                LogError(String.Format("connectTo letter {0}: not allowed", driveLetter));
                Environment.Exit(1);
            }

            int errorCode = Utility.NetworkDrive.MapNetworkDrive(driveLetter, share);
            if (errorCode != 0)
            {
                LogError(String.Format("connectTo MapNetworkDrive exit code = {0}", errorCode));
            }
            Environment.Exit(errorCode);
        }
        */
        static void ConnectToShare(string driveLetter, string share, string shareName)
        {
            if (driveLetter.Length > 1)
            {
                driveLetter = driveLetter.Substring(0, 1);
            }
            driveLetter = driveLetter.ToUpper();
            if (driveLetter.CompareTo("D") == -1 || driveLetter.CompareTo("Z") == 1)
            {
                Utility.Logger.LogError(String.Format("connectTo letter {0}: not allowed", driveLetter));
                Environment.Exit(1);
            }
            
            
            int errorCode = Utility.NetworkDrive.ConnectToShare(driveLetter, share);



            if (errorCode != 0)
            {
                Utility.Logger.LogError(String.Format("connectTo MapNetworkDrive exit code = {0}", errorCode));
                Utility.Logger.LogError(Utility.NetworkDrive.GetErrorMessage(errorCode));
                return;
            }

            if (!string.IsNullOrEmpty(shareName))
            {
                Utility.Logger.LogInformation(string.Format("Set name {0} on {1} for {1}",shareName,driveLetter,share));
                string keyName = share.Replace("\\", "#");
                Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\" + keyName, "_LabelFromDesktopINI", shareName);
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
                ConnectToShare(args[1], args[2], null);
            }
            else if (args.Length == 4 && args[0].Equals("-share"))
            {
                ConnectToShare(args[1], args[3], args[2]);
            }
            else
            {
                PrintUsage();
            }
        }
    }
}
