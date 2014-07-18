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
        public static ActionResult CaActiveSetup_RemoveOpenHKCU(Session session)
        {
            //     <RegistryKey  Root='HKLM' Key='Software\Microsoft\Active Setup\Installed Components\[ProductName]' >
            //  <RegistryValue Action='write' Name='StubPath' Type='string' Value='[CREATE_HKCU_OPEN_CMD]' />
            //</RegistryKey>
            throw new NotImplementedException();
        }

        [CustomAction]
        public static ActionResult CaActiveSetup_SetOpenHKCU(Session session)
        {
            //     <RegistryKey  Root='HKLM' Key='Software\Microsoft\Active Setup\Installed Components\[ProductName]' >
            //  <RegistryValue Action='write' Name='StubPath' Type='string' Value='[CREATE_HKCU_OPEN_CMD]' />
            //</RegistryKey>
            throw new NotImplementedException();
        }

    }
}
