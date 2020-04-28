using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
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

        public static bool ProcessFile(string filename, string checkInComment, ICallback callback = null)
        {
            var ret = false;

            #region event handlers
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
            #endregion

            if (Application == null)
                throw new Exception("Application propertyi is not set.");

            if (callback != null)
                callback.Message = $"Attaching event handlers...";

                var frame = Application.Frame() as Frame;
                var windowHandle = new IntPtr(frame.GetHWndx64());
                var sldworksElement = AutomationElement.FromHandle(windowHandle);
            if (sldworksElement != null)
            {
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, sldworksElement, TreeScope.Children, checkoutDialogOpened);

                var docType = swDocumentTypes_e.swDocPART;
                var extension = System.IO.Path.GetExtension(filename).ToLower();
                if (extension.Contains("sldprt"))
                    docType = swDocumentTypes_e.swDocPART;
                else if (extension.Contains("sldasm"))
                    docType = swDocumentTypes_e.swDocASSEMBLY;
                else if (extension.Contains("slddrw"))
                    docType = swDocumentTypes_e.swDocDRAWING;


                string name = System.IO.Path.GetFileName(filename);

                if (callback != null)
                    callback.Message = $"Opening {name}...";

                var modelDoc = Application.OpenDoc(filename, (int)docType) as ModelDoc2;


                if (modelDoc == null)
                {
                    Application.CloseAllDocuments(true);
                    if (callback != null)
                        callback.Message = $"Failed to open {name}.";
                    Automation.RemoveAllEventHandlers();
                    return ret;
                }
                else
                {
                    if (callback != null)
                        callback.Message = $"Opened {name}.";
                }
                // execute macro

                if (callback != null)
                    callback.Message = $"Attempting to execute macro...";

                int errorMacro = 0;
            

                var macroRet = Application.RunMacro2(MacroParameters.FilePathName, MacroParameters.ModuleName, MacroParameters.ProcedureName, (int)swRunMacroOption_e.swRunMacroDefault, out errorMacro);

                if (callback != null)
                    callback.Message = $"Macro return [{name}] = {macroRet}";

                // if macro does not work not then undo lock
                if (macroRet)
                {
                    int errors = 0;
                    int warnings = 0;
                    var retSave = modelDoc.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);

                    if (callback != null)
                        callback.Message = $"Save return [{name}] = {retSave}";
                    if (retSave)
                    {
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, sldworksElement, TreeScope.Children, checkInDialogOpened);
                        // check in document 
                        addinObject.Unlock();

                    }
                    else
                    {
                        if (callback != null)
                            callback.Message = $"Undoing lock {name}.";
                        addinObject.UndoLock();
                    }


                }
                else
                {
                    if (callback != null)
                        callback.Message = $"Undoing lock {name}.";
                    addinObject.UndoLock();
                }

                // close document
                Application.CloseAllDocuments(true);

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


    public interface ICallback
    {
        string Message { get; set; }
    }

    public struct MacroParameters
    {
        public string FilePathName;
        public string ModuleName;
        public string ProcedureName; 

    }
}
