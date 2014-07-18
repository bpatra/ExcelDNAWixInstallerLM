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
            string szOfficeRegKeyVersions = string.Empty;
            string szBaseAddInKey = @"Software\Microsoft\Office\";
            string szXll32Bit = string.Empty;
            string szXll64Bit = string.Empty;
            string szXllToRegister = string.Empty;
            int nOpenVersion;
            double nVersion;
            bool bFoundOffice = false;
            List<string> lstVersions;

            try
            {
                session.Log("Enter try block of CaRegisterAddIn");

                szOfficeRegKeyVersions = session["OFFICEREGKEYS"];
                szXll32Bit = session["XLL32"];
                szXll64Bit = session["XLL64"];

                if (szOfficeRegKeyVersions.Length > 0)
                {
                    lstVersions = szOfficeRegKeyVersions.Split(',').ToList();

                    foreach (string szOfficeVersionKey in lstVersions)
                    {
                        nVersion = double.Parse(szOfficeVersionKey, NumberStyles.Any, CultureInfo.InvariantCulture);

                        session.Log("Retrieving Registry Information for : " + szBaseAddInKey + szOfficeVersionKey);

                        // get the OPEN keys from the Software\Microsoft\Office\[Version]\Excel\Options key, skip if office version not found.
                        if (Registry.CurrentUser.OpenSubKey(szBaseAddInKey + szOfficeVersionKey, false) != null)
                        {
                            string szKeyName = szBaseAddInKey + szOfficeVersionKey + @"\Excel\Options";

                            szXllToRegister = GetAddInName(szXll32Bit, szXll64Bit, szOfficeVersionKey, nVersion);

                            RegistryKey rkExcelXll = Registry.CurrentUser.OpenSubKey(szKeyName, true);

                            if (rkExcelXll != null)
                            {
                                string[] szValueNames = rkExcelXll.GetValueNames();
                                bool bIsOpen = false;
                                int nMaxOpen = -1;

                                // check every value for OPEN keys
                                foreach (string szValueName in szValueNames)
                                {
                                    // if there are already OPEN keys, determine if our key is installed
                                    if (szValueName.StartsWith("OPEN"))
                                    {
                                        nOpenVersion = int.TryParse(szValueName.Substring(4), out nOpenVersion) ? nOpenVersion : 0;
                                        int nNewOpen = szValueName == "OPEN" ? 0 : nOpenVersion;
                                        if (nNewOpen > nMaxOpen)
                                        {
                                            nMaxOpen = nNewOpen;
                                        }

                                        // if the key is our key, set the open flag
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
                                        rkExcelXll.SetValue("OPEN", "/R \"" + szXllToRegister + "\"");
                                    }
                                    else
                                    {
                                        rkExcelXll.SetValue("OPEN" + (nMaxOpen + 1).ToString(CultureInfo.InvariantCulture), "/R \"" + szXllToRegister + "\"");
                                    }
                                    rkExcelXll.Close();
                                }
                                bFoundOffice = true;
                            }
                            else
                            {
                                session.Log("Unable to retrieve key for : " + szKeyName);
                            }
                        }
                        else
                        {
                            session.Log("Unable to retrieve registry Information for : " + szBaseAddInKey + szOfficeVersionKey);
                        }
                    }
                }

                session.Log("End CaRegisterAddIn");
            }
            catch (System.Security.SecurityException ex)
            {
                session.Log("CaRegisterAddIn SecurityException" + ex.Message);
                bFoundOffice = false;
            }
            catch (System.UnauthorizedAccessException ex)
            {
                session.Log("CaRegisterAddIn UnauthorizedAccessException" + ex.Message);
                bFoundOffice = false;
            }
            catch (Exception ex)
            {
                session.Log("CaRegisterAddIn Exception" + ex.Message);
                bFoundOffice = false;
            }

            return bFoundOffice ? ActionResult.Success : ActionResult.Failure;
        }


        [CustomAction]
        public static ActionResult CaUnRegisterAddIn(Session session)
        {
            string szOfficeRegKeyVersions = string.Empty;
            string szBaseAddInKey = @"Software\Microsoft\Office\";
            string szXll32Bit = string.Empty;
            string szXll64Bit = string.Empty;
            string szXllToUnRegister = string.Empty;
            double nVersion;
            bool bFoundOffice = false;
            List<string> lstVersions;

            try
            {
                session.Log("Begin CaUnRegisterAddIn");

                szOfficeRegKeyVersions = session["OFFICEREGKEYS"];
                szXll32Bit = session["XLL32"];
                szXll64Bit = session["XLL64"];

                if (szOfficeRegKeyVersions.Length > 0)
                {
                    lstVersions = szOfficeRegKeyVersions.Split(',').ToList();

                    foreach (string szOfficeVersionKey in lstVersions)
                    {
                        nVersion = double.Parse(szOfficeVersionKey, NumberStyles.Any, CultureInfo.InvariantCulture);
                        szXllToUnRegister = GetAddInName(szXll32Bit, szXll64Bit, szOfficeVersionKey, nVersion);

                        // only remove keys where office version is found
                        if (Registry.CurrentUser.OpenSubKey(szBaseAddInKey + szOfficeVersionKey, false) != null)
                        {
                            bFoundOffice = true;

                            string szKeyName = szBaseAddInKey + szOfficeVersionKey + @"\Excel\Options";

                            RegistryKey rkAddInKey = Registry.CurrentUser.OpenSubKey(szKeyName, true);
                            if (rkAddInKey != null)
                            {
                                string[] szValueNames = rkAddInKey.GetValueNames();

                                foreach (string szValueName in szValueNames)
                                {
                                    if (szValueName.StartsWith("OPEN") && rkAddInKey.GetValue(szValueName).ToString().Contains(szXllToUnRegister))
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
