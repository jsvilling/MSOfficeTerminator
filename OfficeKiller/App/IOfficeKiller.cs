using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeKiller.App
{
    public delegate void ExecutionDoneHandler(string message);

    public interface IOfficeKiller
    {
        event ExecutionDoneHandler ExecutionDone;

        void KillAll();

        void ChangeConfig(string itemName, bool itemValue);

        Dictionary<string, bool> GetConfig();
    }
}
