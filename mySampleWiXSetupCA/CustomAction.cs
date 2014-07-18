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
            try
            {
                string command = session["REMOVE_HKCU_OPEN_CMD"];
                if (string.IsNullOrEmpty(command))
                {
                    throw new ArgumentException("REMOVE_HKCU_OPEN_CMD is empty");
                }
                session.Log("ActiveSetup Stubpath is: "+command  );
            }
            catch (Exception ex)
            {
                session.Log(ex.Message);
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        [CustomAction]
        public static ActionResult CaActiveSetup_SetOpenHKCU(Session session)
        {
            try
            {
                string command = session["CREATE_HKCU_OPEN_CMD"];
                if (string.IsNullOrEmpty(command))
                {
                    throw new ArgumentException("REMOVE_HKCU_OPEN_CMD is empty");
                }
                session.Log("ActiveSetup Stubpath is: " + command);
            }
            catch (Exception ex)
            {
                session.Log(ex.Message);
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

    }
}
