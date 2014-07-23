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
    /// For the sake of simplicity use only the 32 ActiveSetup registry...
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
            RegistryKey hklmBase = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            return Open( subKey, true, str => hklmBase.OpenSubKey(str, true),str => hklmBase.CreateSubKey(str));
        }

        public RegistryKey OpenOrCreateHkcuKey(string subKey)
        {
            RegistryKey hkcuBase = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
            return Open( subKey, false, str => hkcuBase.OpenSubKey(str, true), str => hkcuBase.CreateSubKey(str));
        }

        public void DeleteHkcuKey(string subKey)
        {
            RegistryKey hkcuBase = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
            _session.Log("Start the delevetion of  HKCU sub key " + subKey);
            hkcuBase.DeleteSubKey(subKey);
        }
    }
}
