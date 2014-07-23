using System.Security;
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
        /// <summary>
        /// Change the active setup responsible for invoking the .exe to put IsInstalled to false and pass the command for removing in the StubPath.
        /// </summary>
        /// <param name="session"><seealso cref="CaParameters"/></param>
        /// <returns>success status</returns>
        [CustomAction]
        public static ActionResult CaActiveSetup_RemoveOpenHKCU(Session session)
        {
            try
            {
                session.Log("Start CaActiveSetup_RemoveOpenHKCU...");
                var registryAbstractor = new RegistryAbstractor(session);

                //Set the Active Setup ot IsInstalled = 0 and change the StubPath
                var caParams = CaParameters.ExtractFromSession(session);
                using(RegistryKey createHklmKey = registryAbstractor.OpenOrCreateHklmKey(caParams.RegistrySubKey))
                {
                    if (createHklmKey != null)
                    {
                        UpdateActiveSetupKey(createHklmKey, caParams.Default, caParams.CreateComponentId,caParams.CreateCommand, caParams.RemoveCommand,caParams.Version,false);
                    }
                }
            }
            catch (SecurityException ex)
            {
                session.Log("CaActiveSetup_RemoveOpenHKCU SecurityException" + ex.Message);
                return ActionResult.Failure;
            }
            catch (UnauthorizedAccessException ex)
            {
                session.Log("CaActiveSetup_RemoveOpenHKCU UnauthorizedAccessException" + ex.Message);
                return ActionResult.Failure;
            }
            catch (Exception ex)
            {
                session.Log("CaActiveSetup_RemoveOpenHKCU Exception" + ex.Message);
                return ActionResult.Failure;
            }
            session.Log("... end CaActiveSetup_RemoveOpenHKCU, successful!");
            return ActionResult.Success;
        }


        /// <summary>
        /// Update/create the active setup responsible for invoking the .exe to put IsInstalled to true and pass the command for installing in the StubPath.
        /// </summary>
        /// <param name="session"><seealso cref="CaParameters"/></param>
        /// <returns>success status</returns>
        [CustomAction]
        public static ActionResult CaActiveSetup_SetOpenHKCU(Session session)
        {
            try
            {
                session.Log("Start CaActiveSetup_SetOpenHKCU...");
                var registryAbstractor = new RegistryAbstractor(session);
                var caParams = CaParameters.ExtractFromSession(session);
             
              
                //install the active setup that will set the OPEN key.
                using (RegistryKey hklmKey = registryAbstractor.OpenOrCreateHklmKey(caParams.RegistrySubKey))
                {
                    UpdateActiveSetupKey(hklmKey, caParams.Default, caParams.CreateComponentId, caParams.CreateCommand, caParams.RemoveCommand, caParams.Version, true);
                }
            }
            catch (SecurityException ex)
            {
                session.Log("CaActiveSetup_SetOpenHKCU SecurityException" + ex.Message);
                return ActionResult.Failure;
            }
            catch (UnauthorizedAccessException ex)
            {
                session.Log("CaActiveSetup_SetOpenHKCU UnauthorizedAccessException" + ex.Message);
                return ActionResult.Failure;
            }
            catch (Exception ex)
            {
                session.Log("CaActiveSetup_SetOpenHKCU Exception" + ex.Message);
                return ActionResult.Failure;
            }
            session.Log("... end CaActiveSetup_SetOpenHKCU, successful!");
            return ActionResult.Success;
        }

        private static void UpdateActiveSetupKey(RegistryKey activeSetupKey, string defaultForKey, string componentId, string commandInstall, string commandUninstall, string version, bool isInstalled)
        {
            if (activeSetupKey == null)
            {
                throw new ArgumentNullException("activeSetupKey");
            }

            activeSetupKey.SetValue("",defaultForKey);
            activeSetupKey.SetValue("ComponentID", componentId, RegistryValueKind.String);
            activeSetupKey.SetValue("StubPath",isInstalled ? commandInstall : commandUninstall, RegistryValueKind.String);

            //Found that . cannot be used for version
            //http://www.sepago.de/e/helge/2010/04/22/active-setup-explained
            activeSetupKey.SetValue("Version",version.Replace('.',','),RegistryValueKind.String); 
            activeSetupKey.SetValue("IsInstalled", isInstalled ? 1 : 0 ,RegistryValueKind.DWord);
        }
    }
}
