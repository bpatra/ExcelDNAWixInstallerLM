using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ActiveSetupCommand
{
    public class Program
    {
        static int Main(string[] args)
        {
            try
            {
                Console.WriteLine("Start extracting args: " + string.Join(";",args));
                var parameters = Parameters.ExtractFromArgs(args);

                switch (parameters.Command)
                {
                    case Command.Install:
                        HkcuKeys.CreateOpenHkcuKey(parameters);
                        break;
                    case Command.Uninstall:
                        HkcuKeys.CreateOpenHkcuKey(parameters);
                        break;
                    default:
                        throw new NotSupportedException("unknown command");
                }
                Console.WriteLine("Command successfully executed!");
                return 0;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Error: "+ exception.Message);
                return 1;
            }    
        }

      
    }
}
