

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace PDMStandardBatchTool.Core
{

    public class Core
    {
        public static MacroParameters MacroParameters;
        public static SldWorks Application { get; set; }

        public static ConisioSW2Lib.IConisioSWAddIn addinObject { get; private set; }

        public static bool ProcessFile(string filename, string checkInComment)
        {
            var ret = false;

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
             


                // add ui automation event handler 
                var frame = Application.Frame() as Frame;
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
          




            return ret;

        }

        public static bool LoadPDMAddIn(string dllPath)
        {
            var ret = false;

            var loadRet = (swLoadAddinError_e)Application.LoadAddIn(dllPath);

            if (loadRet == swLoadAddinError_e.swAddinAlreadyLoaded || loadRet == swLoadAddinError_e.swSuccess)
                ret = true;


                string pdmAddInClsId = Helper.GetPDMAddInGuid();
                addinObject = Application.GetAddInObject(pdmAddInClsId) as ConisioSW2Lib.IConisioSWAddIn;
            if (addinObject == null)
                throw new Exception("Failed to load PDM add-in");

            return ret;
        }

    }

    public struct MacroParameters
    {
        public string FilePathName;
        public string ModuleName;
        public string ProcedureName; 

    }
}
