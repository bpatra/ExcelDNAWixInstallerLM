using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace manageOpenKey
{
    class RegistryAbstractor
    {
        private readonly IWriter _writer;
        public RegistryAbstractor(IWriter writer)
        {
            _writer = writer;
        }
        public RegistryKey OpenOrCreateHkcuKey(string subKey)
        {
            RegistryKey rkExcelXll;
            _writer.WriteLine(string.Format("Opening {0} Key ...", subKey));
            if (Registry.CurrentUser.OpenSubKey(subKey) == null)
            //When triggered by active setup the Excel Options key may not exists create it!
            {
                rkExcelXll = Registry.CurrentUser.CreateSubKey(subKey);
                _writer.WriteLine("... key not existing, create it.");
            }
            else
            {
                rkExcelXll = Registry.CurrentUser.OpenSubKey(subKey, true);
                _writer.WriteLine("... existing key successfully retrieved.");
            }
            return rkExcelXll;
        }
    }
}
