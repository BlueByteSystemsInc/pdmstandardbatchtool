using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using CADSharpTools.SOLIDWORKS;
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
            string checkInComment = "This is a check-in comment";
            var pdmAddInClsId = Helper.GetPDMAddInGuid();
            string fileName = @"C:\PDM2019\PDM2019\30-151127.sldprt";
            string addinDllFilename = @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS PDM\pdmsw.dll";
        
            var checkoutDialogOpened = new AutomationEventHandler((object sender, AutomationEventArgs e) => {

                var window = sender as AutomationElement;

                var windowTitle = window.Current.Name;

                if (windowTitle != "SOLIDWORKS PDM")
                    return;

                bool catchPhraseFound = false; 
                string catchPhrase = string.Empty;
                catchPhrase = "Would you like to check it out?";

                var control = TreeWalker.ControlViewWalker.GetFirstChild(window);
                while (control != null)
                {
                    var text = control.Current.Name;
                    if (string.IsNullOrWhiteSpace(text) == false)
                    {
                        if (text.Contains(catchPhrase))
                            catchPhraseFound = true; 
                    }
                    // to do some stuff here
                    control = TreeWalker.ControlViewWalker.GetNextSibling(control);

                    
                }


                if (catchPhraseFound) 
                { 
                control = TreeWalker.ControlViewWalker.GetFirstChild(window);
                while (control != null)
                {
                    if (control.Current.ControlType == ControlType.Button)
                        {
                            var text = control.Current.Name;
                            if (string.IsNullOrWhiteSpace(text) == false)
                            {
                                if (text.Contains("Yes"))
                                {
                                    var invokePattern = control.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;

                                    if (invokePattern != null)
                                        invokePattern.Invoke();
                                }

                            }

                        }
                   
                    // to do some stuff here
                    control = TreeWalker.ControlViewWalker.GetNextSibling(control);


                }

                }


                Automation.RemoveAllEventHandlers();
            });
            var checkInDialogOpened = new AutomationEventHandler((object sender, AutomationEventArgs e) => {

                var window = sender as AutomationElement;

                var windowTitle = window.Current.Name;

                if (windowTitle != "Check in")
                    return;

               
                var control = TreeWalker.ControlViewWalker.GetFirstChild(window);
             
                while (control != null)
                {
                    if (control.Current.ClassName == "Edit")
                    {

                        control.SetFocus();
                        Thread.Sleep(100);
                        SendKeys.SendWait(checkInComment);
                         

                    }

                    // to do some stuff here
                    control = TreeWalker.ControlViewWalker.GetNextSibling(control);


                }


                control = TreeWalker.ControlViewWalker.GetFirstChild(window);
                    while (control != null)
                    {
                        if (control.Current.ControlType == ControlType.Button)
                        {
                            var text = control.Current.Name;
                            if (string.IsNullOrWhiteSpace(text) == false)
                            {
                                if (text.Contains("Check in"))
                                {
                                    var invokePattern = control.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;

                                    if (invokePattern != null)
                                        invokePattern.Invoke();
                                }

                            }

                        }

                        // to do some stuff here
                        control = TreeWalker.ControlViewWalker.GetNextSibling(control);


                    }

                


                Automation.RemoveAllEventHandlers();
            });

            ConisioSW2Lib.IConisioSWAddIn addinObject = default(ConisioSW2Lib.IConisioSWAddIn);

               var swApp = CADSharpTools.SOLIDWORKS.Application.CreateSldWorks(false, 30);
             
                 
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
                                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, sldworksElement, TreeScope.Children, checkoutDialogOpened);

                            var docType = swDocumentTypes_e.swDocPART;
                            var extension = System.IO.Path.GetExtension(fileName).ToLower();
                            if (extension.Contains("sldprt"))
                                docType = swDocumentTypes_e.swDocPART;
                            else if (extension.Contains("sldasm"))
                                docType = swDocumentTypes_e.swDocASSEMBLY;
                            else if (extension.Contains("slddrw"))
                                docType = swDocumentTypes_e.swDocDRAWING;

                            var modelDoc = swApp.OpenDoc(fileName, (int)docType) as ModelDoc2;


                            //  this is where we would execute a vba macro 
                            modelDoc.Save();
                            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, sldworksElement, TreeScope.Children, checkInDialogOpened);
                            // check in document 
                            addinObject.Unlock();

                            // wait until document is readonly
                            
                            // close document
                            swApp.CloseAllDocuments(true);
                            System.Console.ReadLine();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(swApp);
            

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
