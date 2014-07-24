using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace manageOpenKey
{
    class ConsoleWriter : IWriter
    {
        public ConsoleWriter()
        {
            
        }

        public void Dispose()
        {
            Thread.Sleep(30*1000);
        }

        public void WriteLine(string str, params object[] args)
        {
            Console.WriteLine(str,args);
        }
    }
}
