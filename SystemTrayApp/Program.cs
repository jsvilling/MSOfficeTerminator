using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeTerminator.App;
using OfficeTerminator.Terminators;

namespace SystemTrayApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ApplicationContext applicationContext = new CustomApplicationContext(new OfficeTerminator.Terminators.OfficeTerminator());
            Application.Run(applicationContext);
        }
    }
}
