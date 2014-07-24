using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace manageOpenKey
{
    interface IWriter : IDisposable
    {
        void WriteLine(string str, params object[] args);
    }
}
