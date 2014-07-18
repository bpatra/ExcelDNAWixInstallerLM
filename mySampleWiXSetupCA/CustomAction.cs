using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace mySampleWiXSetupCA
{
    public class CustomActions
    {
        [CustomAction]
        public static ActionResult CaActiveSetupUnregister(Session session)
        {
          throw new NotImplementedException();
        }

    }
}
