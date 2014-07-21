using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;

namespace mySampleWiXSetupCA
{
    /// <summary>
    /// Main parameters that are extracted for <see cref="Session"/> input parameter of the CustomActions
    /// </summary>
    /// <remarks>
    /// Not all parameters are required for both custom actions, however for the sake of simplicity
    /// we pass the same args for <see cref="CustomActions.CaActiveSetup_RemoveOpenHKCU"/> and <see cref="CustomActions.CaActiveSetup_SetOpenHKCU"/>
    /// </remarks>
    public class CaParameters
    {
        const string BasePathActiveSetup = @"SOFTWARE\Microsoft\Active Setup\Installed Components\";

        public string ProductName { get; private set; }
        public string Version { get; private set; }

        /**RELATIVE TO THE CREATION OF OPEN KEY***/
        public string CreateCommand { get; private set; }
        private string CreateActiveSetupGuid { get; set; }
  
        public String CreateRegistrySubKey
        {
            get { return BasePathActiveSetup + CreateActiveSetupGuid; }
        }

        public string CreateComponentId
        {
            get { return ProductName + " ActiveSetup Create Open Key"; }
        }

        /**RELATIVE TO THE REMOVAL OF OPEN KEY***/
        public string RemoveCommand { get; private set; }
        private string RemoveActiveSetupGuid { get; set; }

        public String RemoveRegistrySubKey
        {
            get { return BasePathActiveSetup + RemoveActiveSetupGuid; }
        }

        public string RemoveComponentId
        {
            get { return ProductName + " ActiveSetup Remove Open Key"; }
        }

        public static CaParameters ExtractFromSession(Session session)
        {
            const string createHkcuCommandKey = "CREATE_HKCU_OPEN_FULL";
            const string createGuidKey = "ACTIVESETUP_CREATE_GUID";
            const string productNameKey = "PRODUCTNAME";
            const string versionKey = "VERSION";
            const string removeHkcuCommand = "REMOVE_HKCU_OPEN_FULL";
            const string removeGuidKey = "ACTIVESETUP_REMOVE_GUID";

            string productName = ExtractAndCheck(session, productNameKey);
            string version = ExtractAndCheck(session, versionKey);

            string createCommand = ExtractAndCheck(session, createHkcuCommandKey);
            string createActiveSetupGuid = ExtractAndCheck(session, createGuidKey);

            string removeCommand = ExtractAndCheck(session, removeHkcuCommand);
            string removeActiveSetupGuid = ExtractAndCheck(session, removeGuidKey);

            var caParams = new CaParameters()
                {
                    CreateActiveSetupGuid = createActiveSetupGuid,
                    CreateCommand = createCommand,
                    ProductName = productName,
                    Version = version,
                    RemoveCommand = removeCommand,
                    RemoveActiveSetupGuid = removeActiveSetupGuid
                };

            return caParams;
        }

        private static string ExtractAndCheck(Session session, string key)
        {
            string value = session.CustomActionData[key];
            if (string.IsNullOrEmpty(value))
            {
                throw new ArgumentException(key + "is not found");
            }
            return value;
        }
    }
}
