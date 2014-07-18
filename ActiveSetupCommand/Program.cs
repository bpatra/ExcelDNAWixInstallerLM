using System;
using System.Collections.Generic;
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
                CheckArgs(args);
                Thread.Sleep(5*1000);
                Console.WriteLine("Active setup ran successfully");
                return 0;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine(exception.Message);
                return 1;
            }    
        }

        private static void CheckArgs(string[] args)
        {
            if (args.Length == 0 || args.Length >=2)
            {
                throw new ArgumentException("Wrong number of arguments: one argument possible we found: " + args.Length);
            }
            if (args[0] != @"/i" || args[0] != @"/u")
            {
                throw new ArgumentException(@"There are two arguments possible: /i (install) or /u (uninstall).)");
            }
        }
    }
}
