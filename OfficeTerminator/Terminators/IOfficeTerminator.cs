using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeTerminator.App
{
    public delegate void ExecutionDoneHandler(string message);

    public interface IOfficeTerminator
    {
        event ExecutionDoneHandler AllExecutionDone;

        void TerminateAll();

        void ChangeConfig(string itemName, bool itemValue);

        Dictionary<string, bool> GetConfig();
    }
}
