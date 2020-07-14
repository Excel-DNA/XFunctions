using System.Diagnostics;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;

namespace ExcelDna.XFunctions
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var functions = ExcelRegistration.GetExcelFunctions().ToList();
            if (HasNativeXMatch())
            {
                foreach (var func in functions)
                {
                    func.FunctionAttribute.Name = func.FunctionAttribute.Name + ".FROM.ADDIN";
                }
            }
            functions.RegisterFunctions();

            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        bool HasNativeXMatch()
        {
            int xlfXMatch = 620;
            var retval = XlCall.TryExcel(xlfXMatch, out var _, 1, 1);
            return (retval == XlCall.XlReturn.XlReturnSuccess);
        }
    }
}
