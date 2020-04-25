using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using SolidWorks.Interop.swpublished;

namespace PDMStandardBatchTool
{
    [ComVisible(true)]
    [Guid("CBDB60F7-34E9-4DAF-8A72-44BBD1E21799")]
    public class Main  : SwAddin 
    {
        const string AddInName = "PDMStandardBatchTool";
        const string AddInDescription = "Executes macro on PDM standard files with checkin and check outs.";
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            throw new NotImplementedException();
        }

        public bool DisconnectFromSW()
        {
            throw new NotImplementedException();
        }

        [ComRegisterFunction()]
        public void RegisterAddIn(Type t)
        {
            string KeyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
            RegistryKey rk = Registry.LocalMachine.CreateSubKey(KeyPath);
            rk.SetValue(null, 1); // 1: Add-in will load at start-up
            rk.SetValue("Title", AddInName); // Title
            rk.SetValue("Description", AddInDescription); // Description
        }

        [ComUnregisterFunction()]
        public void UnregisterAddIn(Type t)
        {
            try
            {
                bool Exist = false;
                string KeyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);
                using (RegistryKey Key = Registry.LocalMachine.OpenSubKey(KeyPath))
                {
                    if (Key != null)
                        Exist = true;
                    else
                        Exist = false;
                }
                if (Exist)
                    Registry.LocalMachine.DeleteSubKeyTree(KeyPath);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
