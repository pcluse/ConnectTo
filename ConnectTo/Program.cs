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
    // https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes
    public enum ErrorCodes
    {
        NO_ERROR = 0x0,
        ERROR_ACCESS_DENIED =               0x00000005,
        ERROR_BAD_DEV_TYPE =                0x00000042,
        ERROR_BAD_NET_NAME =                0x00000043,
        ERROR_ALREADY_ASSIGNED =            0x00000055,
        ERROR_INVALID_PASSWORD =            0x00000056,
        ERROR_BUSY =                        0x000000AA,
        ERROR_BAD_DEVICE =                  0x000004B0,
        ERROR_CONNECTION_UNAVAIL =          0x000004B1,
        ERROR_BAD_PROFILE =                 0x000004b6,
        ERROR_NOT_CONNECTED =               0x000008CA,
        ERROR_OPEN_FILES =                  0x00000961,
        ERROR_DEVICE_ALREADY_REMEMBERED =   0x000004B2,
        ERROR_NO_NET_OR_BAD_PATH =          0x000004B3,
        ERROR_CANNOT_OPEN_PROFILE =         0x000004B5,
        ERROR_EXTENDED_ERROR =              0x000004B8,
        ERROR_NO_NETWORK =                  0x000004C6,
        ERROR_CANCELLED =                   0x000004C7
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
            string sLocalName,
            StringBuilder sbRemoteName,
            ref int oilength);

        public static string GetErrorMessage(int ErrorCode)
        {
            //string Message = "";
            switch (ErrorCode)
            {
                case (int)ErrorCodes.ERROR_ACCESS_DENIED:
                    return "Access is denied.";
                case (int)ErrorCodes.ERROR_BAD_DEV_TYPE:
                    return "The network resource type is not correct.";
                case (int)ErrorCodes.ERROR_BAD_NET_NAME:
                    return "The network name cannot be found.";
                case (int)ErrorCodes.ERROR_ALREADY_ASSIGNED:
                    return "The local device name is already in use.";
                case (int)ErrorCodes.ERROR_INVALID_PASSWORD:
                    return "The specified network password is not correct.";
                case (int)ErrorCodes.ERROR_BUSY:
                    return "The requested resource is in use.";
                case (int)ErrorCodes.ERROR_BAD_DEVICE:
                    return "The specified device name is invalid.";
                case (int)ErrorCodes.ERROR_CONNECTION_UNAVAIL:
                    return "The device is not currently connected but it is a remembered connection.";
                case (int)ErrorCodes.ERROR_BAD_PROFILE:
                    return "The network connection profile is corrupted.";
                case (int)ErrorCodes.ERROR_NOT_CONNECTED:
                    return "This network connection does not exist.";
                case (int)ErrorCodes.ERROR_OPEN_FILES:
                    return "This network connection has files open or requests pending.";
                case (int)ErrorCodes.ERROR_DEVICE_ALREADY_REMEMBERED:
                    return "The local device name has a remembered connection to another network resource.";
                case (int)ErrorCodes.ERROR_NO_NET_OR_BAD_PATH:
                    return "The network path was either typed incorrectly, does not exist, or the network provider is not currently available. Please try retyping the path or contact your network administrator.";
                case (int)ErrorCodes.ERROR_CANNOT_OPEN_PROFILE:
                    return "Unable to open the network connection profile.";
                case (int)ErrorCodes.ERROR_EXTENDED_ERROR:
                    return "An extended error has occurred.";
                    /*
                    WNetGetLastErrorA(
                        LPDWORD lpError,
                        LPSTR   lpErrorBuf,
                        DWORD   nErrorBufSize,
                        LPSTR   lpNameBuf,
                        DWORD   nNameBufSize
                        );
                    */
                case (int)ErrorCodes.ERROR_NO_NETWORK:
                    return "The network is not present or not started.";
                case (int)ErrorCodes.ERROR_CANCELLED:
                    return "The operation was canceled by the user.";
                case (int)ErrorCodes.NO_ERROR:
                    return "No error.";
                default:
                    return string.Format("{0}", ErrorCode);
            }
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
        public static int ConnectToShare (string sDriveLetter, string sShare)
        {
            int length = 300;
            int result = 0;
            int timeToWait = 5000;
            int numberOfTries = 10;
            int i = 0;
            StringBuilder currentShare = new StringBuilder(length);
            string lastOperation = "";
            while (!currentShare.ToString().Equals(sShare) && i < numberOfTries)
            {
                // Clear currentShare to avoid logging the wrong path
                currentShare.Clear();
                // Get current mapping
                result = WNetGetConnection((sDriveLetter + ":"), currentShare, ref length);
                lastOperation = "WNetGetConnection";
                Logger.LogInformation(string.Format("Operation {0} gave {1} as result and returned UNC {2}. Message: {3}", lastOperation, result, currentShare.ToString(), GetErrorMessage(result)));
                // If there were no error and the returned share is the same as the user supplied share
                if (result == (int)ErrorCodes.NO_ERROR && currentShare.ToString().Equals(sShare))
                {
                    // Everything is fine, return
                    Logger.LogInformation(string.Format("{1}: is connected to {1}",sDriveLetter,sShare));
                    return result;
                }
                /*
                 * Cancelling a disconnected persistent connection results in ERROR_CONNECTION_UNAVAIL
                 * the connection needs to be restored before it can be cancelled
                 */
                else if (result == (int)ErrorCodes.ERROR_CONNECTION_UNAVAIL)
                {
                    // Create a NETRESOURCE which corresponds to the current input
                    NETRESOURCE oNetworkResource = new NETRESOURCE()
                    {
                        oResourceType = ResourceType.RESOURCETYPE_DISK,
                        sLocalName = sDriveLetter + ":",
                        sRemoteName = currentShare.ToString()
                    };
                    // Add a persistent connection
                    result = WNetAddConnection2(ref oNetworkResource, null, null, 1);
                    lastOperation = "WNetAddConnection2 - existing connection";
                }
                // Not connected, connect it!
                else if (result == (int)ErrorCodes.ERROR_NOT_CONNECTED)
                {
                    // Create a NETRESOURCE which corresponds to the current input
                    NETRESOURCE oNetworkResource = new NETRESOURCE()
                    {
                        oResourceType = ResourceType.RESOURCETYPE_DISK,
                        sLocalName = sDriveLetter + ":",
                        sRemoteName = sShare
                    };
                    // Add a persistent connection
                    result = WNetAddConnection2(ref oNetworkResource, null, null, 1);
                    lastOperation = "WNetAddConnection2";
                }
                /* 
                    NO_ERROR means it retrieved the connection successfully. 
                    If it does not match the path it should cancel the existing connection.
                */
                else if (result == (int)ErrorCodes.NO_ERROR && !currentShare.ToString().Equals(sShare))
                {
                    // Disconnect (cancel) the current connection
                    result = WNetCancelConnection2(sDriveLetter + ":", 1, 1);
                    lastOperation = "WNetCancelConnection2";
                }

                Logger.LogInformation(string.Format("Operation {0} gave {1} as result. Message: {2}",lastOperation,result,GetErrorMessage(result)));
                // To avoid hammering, wait for a while
                Thread.Sleep(timeToWait);
                i++;
            }
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

        static void ConnectToShare(string driveLetter, string share, string shareName)
        {
            Utility.Logger.LogInformation("ConnectTo -share start.");
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
                Utility.Logger.LogError(String.Format("connectTo MapNetworkDrive exit code = {0}: {1}", errorCode, Utility.NetworkDrive.GetErrorMessage(errorCode)));
                return;
            }

            if (!string.IsNullOrEmpty(shareName))
            {
                Utility.Logger.LogInformation(string.Format("Set name '{0}' on {1} for {2}",shareName,driveLetter,share));
                string keyName = share.Replace("\\", "#");
                Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\" + keyName, "_LabelFromDesktopINI", shareName);
            }
            Utility.Logger.LogInformation("ConnectTo -share is done.");
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
