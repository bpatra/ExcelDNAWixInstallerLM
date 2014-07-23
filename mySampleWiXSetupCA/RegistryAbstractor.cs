using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;

namespace mySampleWiXSetupCA
{
    /// <summary>
    /// Wraps registry action for logging.
    /// </summary>
    /// <remarks>
    /// This may or may not write in the Wow6432 but actually we do not care....
    /// </remarks>
    class RegistryAbstractor
    {
        private readonly Session _session;
        public RegistryAbstractor(Session session)
        {
            _session = session;
        }

        private RegistryKey Open(string subkey, bool isHklm, Func<string,RegistryKey> providerOpen, Func<string,RegistryKey> providerCreate)
        {
            _session.Log(string.Format("Opening {0} Key {1}...", isHklm ? "HKLM " : "HKCU", subkey));
            RegistryKey registryKey = providerOpen(subkey);

            if (registryKey == null)
            {
                _session.Log("... key not existing, create it.");
                registryKey = providerCreate(subkey);
            }
            else
            {
                _session.Log("... existing key successfully retrieved.");
            }

            return registryKey;
        }

        public RegistryKey OpenOrCreateHklmKey(string subKey)
        {
            return Open( subKey, true, str => Registry.LocalMachine.OpenSubKey(str, true),
                        str => Registry.LocalMachine.CreateSubKey(str));
        }

        public RegistryKey OpenOrCreateHkcuKey(string subKey)
        {
            return Open( subKey, false, str => Registry.CurrentUser.OpenSubKey(str, true),
                        str => Registry.CurrentUser.CreateSubKey(str));
        }

        public void DeleteHkcuKey(string subKey)
        {
            _session.Log("Start the delevetion of  HKCU sub key " + subKey);
            Registry.CurrentUser.DeleteSubKey(subKey);
        }
    }
}
