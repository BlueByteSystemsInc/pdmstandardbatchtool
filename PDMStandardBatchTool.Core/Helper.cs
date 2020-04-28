namespace PDMStandardBatchTool.Core
{
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
