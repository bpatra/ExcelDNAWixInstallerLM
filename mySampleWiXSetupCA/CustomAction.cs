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
        /// 1) uninstall the active setup responsible for invoking the .exe (that will set the OPEN HKCU key for each user).
        /// 2) install the active responsible for invoking the .exe (that will remove the OPEN HKCU key for each user).
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

                //1) set the Active Setup (HKLM and current user) that sets the OPEN Key to installed : 0
                var caParams = CaParameters.ExtractFromSession(session);
                using(RegistryKey createHklmKey = Registry.LocalMachine.OpenSubKey(caParams.CreateRegistrySubKey, true))
                {
                    if (createHklmKey != null)
                    {
                        UpdateActiveSetupKey(createHklmKey, caParams.DefaultCreate, caParams.CreateComponentId,caParams.CreateCommand,caParams.Version,false);
                    }
                }
             

                //PRO Tip: create the HKCU active setup key directly to avoid to trigger active setup when the current user will logon...
                using (RegistryKey createHkcuKey = Registry.CurrentUser.OpenSubKey(caParams.CreateRegistrySubKey, true))
                {
                    if (createHkcuKey != null)
                    {
                        UpdateActiveSetupKey(createHkcuKey, caParams.DefaultCreate, caParams.CreateComponentId, caParams.CreateCommand, caParams.Version, false);
                    }
                }
             

                
                //2)set the Active Setup that will remove the Open key.
                using(RegistryKey removeHklmKey = registryAbstractor.OpenOrCreateHklmKey(caParams.RemoveRegistrySubKey))
                {
                    UpdateActiveSetupKey(removeHklmKey, caParams.DefaultRemove, caParams.RemoveComponentId, caParams.RemoveCommand, caParams.Version, true);
                }
                

                //PRO Tip: create the HKCU active setup key directly to avoid to trigger active setup when the current user will logon...
                using (RegistryKey removeHkcuKey = registryAbstractor.OpenOrCreateHkcuKey(caParams.RemoveRegistrySubKey))
                {
                    UpdateActiveSetupKey(removeHkcuKey, caParams.DefaultRemove, caParams.RemoveComponentId, caParams.RemoveCommand, caParams.Version, true);
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
        /// 1) uninstall the active setup responsible for invoking the .exe (that will remove the OPEN HKCU key for each user). Remind that the addin may have already been uninstalled.
        /// 2) install the active setup responsible for invoking the .exe (that will set the OPEN HKCU key for each user).
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
                //1) uninstall the active setup for removing the keys that could have been set before.
                using (RegistryKey removeHklmKey = Registry.LocalMachine.OpenSubKey(caParams.RemoveRegistrySubKey, true))
                {
                    if (removeHklmKey != null)
                    {
                        UpdateActiveSetupKey(removeHklmKey, caParams.DefaultRemove, caParams.RemoveComponentId, caParams.RemoveCommand, caParams.Version, false);
                    }
                }
             
                //PRO Tip: create the HKCU active setup key directly to avoid to trigger active setup when the current user will logon...
                using(RegistryKey removeHkcuKey = Registry.CurrentUser.OpenSubKey(caParams.RemoveRegistrySubKey, true))
                {
                    if (removeHkcuKey != null)
                    {
                        UpdateActiveSetupKey(removeHkcuKey, caParams.DefaultRemove, caParams.RemoveCommand, caParams.RemoveCommand, caParams.Version, false);
                    }
                }
              
                //2) install the active setup that will set the OPEN key.
                using (RegistryKey hklmKey = registryAbstractor.OpenOrCreateHklmKey(caParams.CreateRegistrySubKey))
                {
                    UpdateActiveSetupKey(hklmKey, caParams.DefaultCreate, caParams.CreateComponentId, caParams.CreateCommand, caParams.Version, true);
                }

                //PRO Tip: create the HKCU active setup key directly to avoid to trigger active setup when the current user will logon...
                using (RegistryKey hkcuKey = registryAbstractor.OpenOrCreateHkcuKey(caParams.CreateRegistrySubKey))
                {
                    UpdateActiveSetupKey(hkcuKey, caParams.DefaultCreate, caParams.CreateComponentId, caParams.CreateCommand, caParams.Version, true);
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

        private static void UpdateActiveSetupKey(RegistryKey activeSetupKey, string defaultForKey, string componentId, string command, string version, bool isInstalled)
        {
            if (activeSetupKey == null)
            {
                throw new ArgumentNullException("activeSetupKey");
            }

            activeSetupKey.SetValue("",defaultForKey);
            activeSetupKey.SetValue("ComponentID", componentId, RegistryValueKind.String);
            activeSetupKey.SetValue("StubPath", command, RegistryValueKind.String);

            //Found that . cannot be used for version
            //http://www.sepago.de/e/helge/2010/04/22/active-setup-explained
            activeSetupKey.SetValue("Version",version.Replace('.',','),RegistryValueKind.String); 
            activeSetupKey.SetValue("IsInstalled", isInstalled ? 1 : 0 ,RegistryValueKind.DWord);
        }
    }
}
