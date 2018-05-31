using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Runtime.InteropServices;


/* Reference

    https://stackoverflow.com/questions/43173970/map-network-drive-programmatically-in-c-sharp-on-windows-10?rq=1

    https://stackoverflow.com/questions/16040148/add-printer-to-local-computer-using-managementclass
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

        /*
        public static bool IsDriveMapped(string sDriveLetter)
        {
            string[] DriveList = Environment.GetLogicalDrives();
            for (int i = 0; i < DriveList.Length; i++)
            {
                if (sDriveLetter + ":\\" == DriveList[i].ToString())
                {
                    return true;
                }
            }
            return false;
        } */

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

}


namespace ConnectTo
{
    class Program
    {
        static void PrintUsage()
        {
            Console.WriteLine("Usage: ConnectTo (-printer|-share letter) resource");
            Environment.Exit(1);
        }

        static void ConnectToPrinter(string printer)
        {
            uint errorCode = 0;
            // Console.WriteLine("Connecting to " + printer);
            using (ManagementClass win32Printer = new ManagementClass("Win32_Printer"))
            {
                using (ManagementBaseObject inputParam =
                   win32Printer.GetMethodParameters("AddPrinterConnection"))
                {
                    // Replace <server_name> and <printer_name> with the actual server and
                    // printer names.
                    inputParam.SetPropertyValue("Name", printer);

                    using (ManagementBaseObject result =
                        (ManagementBaseObject)win32Printer.InvokeMethod("AddPrinterConnection", inputParam, null))
                    {
                        errorCode = (uint)result.Properties["returnValue"].Value;

                    }
                }
            }
            Environment.Exit((int)errorCode);
        }

        static void ConnectToShare(string driveLetter, string share)
        {
            // Console.WriteLine("Connecting " + driveLetter + " to " + share);
            if (driveLetter.Length > 1)
            {
                driveLetter = driveLetter.Substring(0, 1);
            }

            int errorCode = Utility.NetworkDrive.MapNetworkDrive(driveLetter, share);
            Environment.Exit(errorCode);
        }

        static void Main(string[] args)
        {
            if (args.Length == 2 && args[0].Equals("-printer"))
            {
                ConnectToPrinter(args[1]);
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
