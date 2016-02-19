using System;
using ExcelDna.Integration;
using ExcelDna.Logging;

namespace Ribbon2
{
    public class Program : IExcelAddIn
    {
        public void AutoOpen()
        {
            try
            {
                
            }
            catch (Exception e)
            {
                LogDisplay.RecordLine(e.Message);
                LogDisplay.RecordLine(e.StackTrace);
                LogDisplay.Show();
            }
        }

        public void AutoClose()
        {
            throw new System.NotImplementedException();
        }

    }
}
