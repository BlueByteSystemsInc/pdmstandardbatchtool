using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using ConisioSW2Lib;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Shell;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace PDMStandardBatchTool.Console
{
    public class EntryPoint
    {

        public static void Main()
        {
            var pdmAddInClsId = Helper.GetPDMAddInGuid();
            var windowOpenedEventHandler = new AutomationEventHandler((object sender, AutomationEventArgs e) => {

                var window = sender as AutomationElement;
                var control = TreeWalker.ControlViewWalker.GetFirstChild(window);
                while (control != null)
                {
                    // to do some stuff here




                    control = TreeWalker.ControlViewWalker.GetNextSibling(control);
                }

            });


            string addinDllFilename = string.Empty;
            ConisioSW2Lib.IConisioSWAddIn addinObject = default(ConisioSW2Lib.IConisioSWAddIn);

            using (var app = new CADSharpTools.SOLIDWORKS.Application(-1, true))
            {
                var swApp = app.SOLIDWORKS;

                addinDllFilename = @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS PDM\pdmsw.dll";

                var loadRet = (swLoadAddinError_e)swApp.LoadAddIn(addinDllFilename);

                switch (loadRet)
                {
                    case swLoadAddinError_e.swUnknownError:
                        break;
                    case swLoadAddinError_e.swSuccess:
                        // get addin object
                        pdmAddInClsId = Helper.GetPDMAddInGuid();
                        addinObject = swApp.GetAddInObject(pdmAddInClsId) as ConisioSW2Lib.IConisioSWAddIn;
                        if (addinObject == null)
                            System.Console.WriteLine("Addin object is null");
                        else
                        {
                            System.Console.WriteLine("Addin object is not null");

                           
                            // add ui automation event handler 
                            var frame = swApp.Frame() as Frame;
                            var windowHandle = new IntPtr(frame.GetHWndx64());
                            var sldworksElement = AutomationElement.FromHandle(windowHandle);
                            if (sldworksElement != null)
                            {
                                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, sldworksElement, TreeScope.Subtree, windowOpenedEventHandler);
                            }
                        
                        }
                            break;
                    #region cases to handle later 
                    case swLoadAddinError_e.swAddinNotLoaded:
                        break;
                    case swLoadAddinError_e.swAddinAlreadyLoaded:
                        break;
                    case swLoadAddinError_e.swFileNotFound:
                        break;
                    case swLoadAddinError_e.swAddinsDisabled:
                        break;
                    case swLoadAddinError_e.swLoadConflict:
                        break;
                    case swLoadAddinError_e.swRegistrationError:
                        break;
                    case swLoadAddinError_e.swLicenseError:
                        break;
                    default:
                        break;

                        #endregion 
                }

                

                swApp.ExitApp();
            }
            

            System.Console.ReadLine();
        }
    }

   public static class Helper
    {
        public static string GetPDMAddInGuid()
        {
            const string PDMAddInName = "SOLIDWORKS PDM";
            var keys = Microsoft.Win32.Registry.LocalMachine.OpenSubKey($@"SOFTWARE\SolidWorks\AddIns").GetSubKeyNames();

            if (keys != null)
            {
                foreach (var key in keys)
                {
                    var title = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\\SolidWorks\AddIns\{key}").GetValue("Title").ToString();

                    if (title == PDMAddInName)
                        return key;
                }
            }

            return string.Empty;

        }

       
    }
}
