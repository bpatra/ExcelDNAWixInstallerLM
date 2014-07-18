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
        public static ActionResult CaRegisterAddIn(Session session)
        {

            string szOfficeRegKeyVersions = session["OFFICEREGKEYS_PROP"];
            string szXll32Bit = session["XLL32_PROP"];
            string szXll64Bit = session["XLL64_PROP"];
            string installFolder = session["INSTALLFOLDER"];

            bool bFoundOffice = CreateOpenHKCUKey(new SessionLogger(session), szOfficeRegKeyVersions, szXll32Bit, szXll64Bit, installFolder);

            return bFoundOffice ? ActionResult.Success : ActionResult.Failure;
        }

        private static bool CreateOpenHKCUKey(ILogger logger, string szOfficeRegKeyVersions, string szXll32Bit, string szXll64Bit, string installFolder)
        {
            bool success = false;
            try
            {
                logger.Log("Enter try block of CreateOpenHKCUKey");

                logger.Log(string.Format("szOfficeRegKeyVersions:{0};szXll32Bit:{1};szXll64Bit:{2};installFolder:{3}",
                                          szOfficeRegKeyVersions, szXll32Bit, szXll64Bit, installFolder));

                if (szOfficeRegKeyVersions.Length > 0)
                {
                    List<string> lstVersions = szOfficeRegKeyVersions.Split(',').ToList();

                    foreach (string szOfficeVersionKey in lstVersions)
                    {
                        double nVersion = double.Parse(szOfficeVersionKey, NumberStyles.Any, CultureInfo.InvariantCulture);

                        logger.Log("Retrieving Registry Information for : " + Constants.SzBaseAddInKey + szOfficeVersionKey);

                        // get the OPEN keys from the Software\Microsoft\Office\[Version]\Excel\Options key, skip if office version not found.
                        if (Registry.CurrentUser.OpenSubKey(Constants.SzBaseAddInKey + szOfficeVersionKey, false) != null)
                        {
                            string szKeyName = Constants.SzBaseAddInKey + szOfficeVersionKey + @"\Excel\Options";

                            string szXllToRegister = GetAddInName(szXll32Bit, szXll64Bit, szOfficeVersionKey, nVersion);
                            //for a localmachine install the xll's should be in the installFolder
                            string fullPathToXll = Path.Combine(installFolder, szXllToRegister);

                            RegistryKey rkExcelXll = Registry.CurrentUser.OpenSubKey(szKeyName, true);
                            if (rkExcelXll != null)
                            {
                                logger.Log("Success finding HKCU key for : " + szKeyName);
                                string[] szValueNames = rkExcelXll.GetValueNames();
                                bool bIsOpen = false;
                                int nMaxOpen = -1;

                                // check every value for OPEN keys
                                foreach (string szValueName in szValueNames)
                                {
                                    // if there are already OPEN keys, determine if our key is installed
                                    if (szValueName.StartsWith("OPEN"))
                                    {
                                        int nOpenVersion = int.TryParse(szValueName.Substring(4), out nOpenVersion)
                                                               ? nOpenVersion
                                                               : 0;
                                        int nNewOpen = szValueName == "OPEN" ? 0 : nOpenVersion;
                                        if (nNewOpen > nMaxOpen)
                                        {
                                            nMaxOpen = nNewOpen;
                                        }

                                        // if the key is our key, set the open flag
                                        //NOTE: this line means if the user has changed its office from 32 to 64 (or conversly) without removing the addin then we will not update the key properly
                                        //The user will have to uninstall addin before installing it again
                                        if (rkExcelXll.GetValue(szValueName).ToString().Contains(szXllToRegister))
                                        {
                                            bIsOpen = true;
                                        }
                                    }
                                }

                                // if adding a new key
                                if (!bIsOpen)
                                {
                                    if (nMaxOpen == -1)
                                    {
                                        rkExcelXll.SetValue("OPEN", "/R \"" + fullPathToXll + "\"");
                                    }
                                    else
                                    {
                                        rkExcelXll.SetValue("OPEN" + (nMaxOpen + 1).ToString(CultureInfo.InvariantCulture),
                                                            "/R \"" + fullPathToXll + "\"");
                                    }
                                    rkExcelXll.Close();
                                }
                                success = true;
                            }
                            else
                            {
                                logger.Log("Unable to retrieve key for : " + szKeyName);
                            }
                        }
                        else
                        {
                            logger.Log("Unable to retrieve registry Information for : " + Constants.SzBaseAddInKey +
                                        szOfficeVersionKey);
                        }
                    }
                }
                else
                {
                    logger.Log("The list of supported office version is empty");
                }

                logger.Log("End CreateOpenHKCUKey");
            }
            catch (System.Security.SecurityException ex)
            {
                logger.Log("CreateOpenHKCUKey SecurityException" + ex.Message);
                success = false;
            }
            catch (UnauthorizedAccessException ex)
            {
                logger.Log("CreateOpenHKCUKey UnauthorizedAccessException" + ex.Message);
                success = false;
            }
            catch (Exception ex)
            {
                logger.Log("CreateOpenHKCUKey Exception" + ex.Message);
                success = false;
            }
            return success;
        }


        [CustomAction]
        public static ActionResult CaUnRegisterAddIn(Session session)
        {
            bool bFoundOffice = false;
            try
            {
                session.Log("Begin CaUnRegisterAddIn");

                string szOfficeRegKeyVersions = session["OFFICEREGKEYS_PROP"];
                string szXll32Bit = session["XLL32_PROP"];
                string szXll64Bit = session["XLL64_PROP"];
                string installFolder = session["INSTALLFOLDER"];

                session.Log(string.Format("szOfficeRegKeyVersions:{0};szXll32Bit:{1};szXll64Bit:{2};installFolder:{3}", szOfficeRegKeyVersions, szXll32Bit, szXll64Bit, installFolder));

                if (szOfficeRegKeyVersions.Length > 0)
                {
                    List<string> lstVersions = szOfficeRegKeyVersions.Split(',').ToList();

                    foreach (string szOfficeVersionKey in lstVersions)
                    {
                        // only remove keys where office version is found
                        if (Registry.CurrentUser.OpenSubKey(Constants.SzBaseAddInKey + szOfficeVersionKey, false) != null)
                        {
                            bFoundOffice = true;

                            string szKeyName = Constants.SzBaseAddInKey + szOfficeVersionKey + @"\Excel\Options";

                            RegistryKey rkAddInKey = Registry.CurrentUser.OpenSubKey(szKeyName, true);
                            if (rkAddInKey != null)
                            {
                                string[] szValueNames = rkAddInKey.GetValueNames();

                                foreach (string szValueName in szValueNames)
                                {
                                    //unregister both 32 and 64 xll
                                    if (szValueName.StartsWith("OPEN") && (rkAddInKey.GetValue(szValueName).ToString().Contains(szXll32Bit) || rkAddInKey.GetValue(szValueName).ToString().Contains(szXll64Bit)))
                                    {
                                        rkAddInKey.DeleteValue(szValueName);
                                    }
                                }
                            }
                        }
                    }
                }

                session.Log("End CaUnRegisterAddIn");
            }
            catch (Exception ex)
            {
                session.Log(ex.Message);
            }

            return bFoundOffice ? ActionResult.Success : ActionResult.Failure;
        }

        //Using a registry key of outlook to determine the bitness of office may look like weird but that's the reality.
        //http://stackoverflow.com/questions/2203980/detect-whether-office-2010-is-32bit-or-64bit-via-the-registry
        public static string GetAddInName(string szXll32Name, string szXll64Name, string szOfficeVersionKey, double nVersion)
        {
            string szXllToRegister = string.Empty;

            if (nVersion >= 14)
            {
                // determine if office is 32-bit or 64-bit
                RegistryKey rkBitness = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Office\" + szOfficeVersionKey + @"\Outlook", false);
                if (rkBitness != null)
                {
                    object oBitValue = rkBitness.GetValue("Bitness");
                    if (oBitValue != null)
                    {
                        if (oBitValue.ToString() == "x64")
                        {
                            szXllToRegister = szXll64Name;
                        }
                        else
                        {
                            szXllToRegister = szXll32Name;
                        }
                    }
                    else
                    {
                        szXllToRegister = szXll32Name;
                    }
                }
            }
            else
            {
                szXllToRegister = szXll32Name;
            }

            return szXllToRegister;
        }

    }
}
