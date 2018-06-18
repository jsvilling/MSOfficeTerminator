//using Microsoft.Office.Interop.Excel;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace OfficeKiller.Killers.OfficeApplicationKiller
//{
//    public class ExcelKiller : OfficeApplicationKiller<Application>
//    {
//        public override string ProcessName => "Excel.Application";

//        public override void Kill()
//        {
//            Application runningExcelApp = FindRunningInstance();
//            if (runningExcelApp != null)
//            {
//                KillExcel(runningExcelApp);
//            }
//        }

//        private void KillExcel(Application runningExcelApp)
//        {
//            foreach (Workbook workBook in runningExcelApp.Workbooks)
//            {
//                workBook.Save();
//                workBook.Close();
//            }
//            runningExcelApp.Quit();
//        }
//    }
//}
