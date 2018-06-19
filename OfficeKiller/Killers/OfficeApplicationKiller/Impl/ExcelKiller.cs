using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller.Killers.OfficeApplicationKiller
{
    public class ExcelKiller : OfficeApplicationKiller<Application>
    {
        public override string ProcessName => "EXCEL";

        public override string InstanceName => "Excel.Application";

        protected override void SaveAndQuit(Application appInstance)
        {
            foreach (Workbook workBook in appInstance.Workbooks)
            {
                workBook.Save();
                workBook.Close();
            }
            appInstance.Quit();
        }
    }
}
