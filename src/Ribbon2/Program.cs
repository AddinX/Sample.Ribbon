using ExcelDna.Integration;
using NetOffice.ExcelApi;

namespace Ribbon2
{
    public class Program : IExcelAddIn
    {
        public void AutoOpen()
        {
            // The Excel Application object
            AddinContext.ExcelApp = new Application(null, ExcelDnaUtil.Application);
        }

        public void AutoClose()
        {
            throw new System.NotImplementedException();
        }

    }
}
