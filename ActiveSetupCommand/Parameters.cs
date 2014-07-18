using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ActiveSetupCommand
{
    class Parameters
    {
        public Command Command { get; set; }
        public string XllName { get; set; }
        public string Xll64Name { get; set; }
        public string InstallDirectory { get; set; }
        public List<string> SupportedOfficeVersion { get; set; } 

          // expected syntax should be
        // pathToExe.exe /install mysample-AddIn-packed.xll mysample-AddIn64-packed.xll 'C:\Program Files\SampleCompany\MySample' '12.0,14.0,15.0'
        // pathToExe.exe /uninstall mysample-AddIn-packed.xll mysample-AddIn64-packed.xll 'C:\Program Files\SampleCompany\MySample' '12.0,14.0,15.0'
        public static Parameters ExtractFromArgs(string[] args)
        {
            if (args.Length !=5)
            {
                throw new ArgumentException("Wrong number of arguments, should be 5 we found: " + args.Length);
            }
            var parameters = new Parameters();
            if (args[0] == "/install")
            {
                parameters.Command = Command.Install;
            }
            else if (args[0] == "/install")
            {
                parameters.Command = Command.Uninstall;
            }
            else
            {
                throw new ArgumentException(@"There are two arguments possible: /install or /uninstall).");
            }

            parameters.XllName = args[1];
            parameters.Xll64Name = args[2];

            if (!Directory.Exists(args[3]))
            {
                throw new DirectoryNotFoundException("Directory not found: " +args[3]);
            }
            parameters.InstallDirectory = args[3];

            parameters.SupportedOfficeVersion = args[4].Split(',').ToList();

            return parameters;
        }

    }

    enum Command
    {
        Install,Uninstall
    }
}
