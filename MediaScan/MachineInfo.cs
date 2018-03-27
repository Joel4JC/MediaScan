using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.NetworkInformation;
using System.Security.Principal;

namespace MediaScan
{
    class MachineInfo
    {
        private string strMacAddress;
        private string strMachineName;
        private string strLoggedInUser;
        private string strUserEnvironment;
        private string strHostName;
        private string strLocalIPAddress;

        /********************************************************************************************
         * These are the accessors for the following public properties. They are used to get the
         * value of the property.
         * 
         ********************************************************************************************/
        public string MacAddress
        {
            get
            {
                return strMacAddress;
            }
        }

        public string MachineName
        {
            get
            {
                return strMachineName;
            }
        }

        public string LoggedInUser
        {
            get
            {
                return strLoggedInUser;
            }
        }

        public string UserEnvironment
        {
            get
            {
                return UserEnvironment;
            }
        }

        public string HostName
        {
            get
            {
                return strHostName;
            }
        }

        public string LocalIPAddress
        {
            get
            {
                return strLocalIPAddress;
            }
        }
        /* End of the Accessors */


        /********************************************************************************************
         * MachineInfo
         * 
         * MachineInfo Constructor, which Calls the various function to gather the various machine 
         * information.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        public MachineInfo()
        {
            GetMACAddress();
            GetMachineName();
            GetUserName();
            GetUserEnviron();
            GetLocalIPAddress();
        } // End of Contructor MachineInfo


        /********************************************************************************************
         * GetMACAddress
         * 
         * Gets the NIC MAC Address for each NIC card in the machine.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void GetMACAddress()
        {
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();

            String sMacAddress = string.Empty;

            foreach (NetworkInterface adapter in nics)
            {
                if (sMacAddress == String.Empty)// only return MAC Address from first card  
                {
                    IPInterfaceProperties properties = adapter.GetIPProperties();

                    sMacAddress = adapter.GetPhysicalAddress().ToString();
                }
            }
            string MA = sMacAddress;
            for (int iNum = 2; iNum < 16; iNum += 3)
                MA = MA.Insert(iNum, ":");

            strMacAddress = MA.ToString();
        } // End of GetMACAddress


        /********************************************************************************************
         * GetMachineName
         * 
         * Gets the name of the computer from the Environment
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void GetMachineName()
        {

            strMachineName = Environment.MachineName;
        } // End of GetMachineName


        /********************************************************************************************
         * GetUsername
         * 
         * Gets the current user's name (the identity under which the thread is running).
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void GetUserName()
        {
            strLoggedInUser = WindowsIdentity.GetCurrent().Name.ToString();
        } // End of GetUsername


        /********************************************************************************************
         * GetuserEnviron
         * 
         * Gets the user name of the person who is currently logged on to the machine.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void GetUserEnviron()
        {
            strUserEnvironment = Environment.UserName;
        } // End of GetUserEnviron


        /********************************************************************************************
         * GetLocalIPAddress
         * 
         * Gets the IP Address of teh machine.
         * 
         * Returns: Nothing
         * 
         ********************************************************************************************/
        private void GetLocalIPAddress()
        {
            IPHostEntry myHost;
            string IP = "?";
            myHost = Dns.GetHostEntry(Dns.GetHostName());

            foreach (IPAddress IPA in myHost.AddressList)
            {
                if (IPA.AddressFamily.ToString() == "InterNetwork")
                {
                    IP = IPA.ToString();
                }
            }

            // If this code is made a separate function, 
            // the values below will have to be returned.
            strHostName = Dns.GetHostName();
            strLocalIPAddress = IP;
        } // End of GetLocalIPAddress
    }
}
