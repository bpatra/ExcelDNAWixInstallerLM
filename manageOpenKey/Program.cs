using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace manageOpenKey
{
    public class Program
    {
        static int Main(string[] args)
        {
            using (IWriter writer = new ConsoleWriter())
            {
                var keyManager = new HkcuKeys(writer);
                try
                {
                    writer.WriteLine("Start extracting args: " + string.Join(";", args));
                    var parameters = Parameters.ExtractFromArgs(args);

                    switch (parameters.Command)
                    {
                        case Command.Install:
                            keyManager.CreateOpenHkcuKey(parameters);
                            break;
                        case Command.Uninstall:
                            keyManager.RemoveHkcuOpenKey(parameters);
                            break;
                        default:
                            throw new NotSupportedException("unknown command");
                    }
                    writer.WriteLine("Command successfully executed!");
                    return 0;
                }
                catch (Exception exception)
                {
                    writer.WriteLine("Error: " + exception.Message);
                    return 1;
                }
            }

        }


    }
}
