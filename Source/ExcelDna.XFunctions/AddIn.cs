using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ExcelDna.XFunctions
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
