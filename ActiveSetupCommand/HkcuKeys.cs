using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace ActiveSetupCommand
{
    public static class HkcuKeys
    {
        public const string SzBaseAddInKey = @"Software\Microsoft\Office\";

        public static bool CreateOpenHKCUKey(string officeRegKeyVersions, string xll32Name, string xll64Name, string fullPathInstallFolder)
        {
            bool success = false;
            try
            {
                Console.WriteLine("Enter try block of CreateOpenHKCUKey");

                Console.WriteLine(string.Format("officeRegKeyVersions:{0};xll32Name:{1};xll64Name:{2};fullPathInstallFolder:{3}",
                                          officeRegKeyVersions, xll32Name, xll64Name, fullPathInstallFolder));

                if (officeRegKeyVersions.Length > 0)
                {
                    List<string> lstVersions = officeRegKeyVersions.Split(',').ToList();

                    foreach (string szOfficeVersionKey in lstVersions)
                    {
                        double nVersion = double.Parse(szOfficeVersionKey, NumberStyles.Any, CultureInfo.InvariantCulture);

                        Console.WriteLine("Retrieving Registry Information for : " + SzBaseAddInKey + szOfficeVersionKey);

                        // get the OPEN keys from the Software\Microsoft\Office\[Version]\Excel\Options key, skip if office version not found.
                        if (Registry.CurrentUser.OpenSubKey(SzBaseAddInKey + szOfficeVersionKey, false) != null)
                        {
                            string szKeyName = SzBaseAddInKey + szOfficeVersionKey + @"\Excel\Options";

                            string szXllToRegister = GetAddInName(xll32Name, xll64Name, szOfficeVersionKey, nVersion);
                            //for a localmachine install the xll's should be in the installFolder
                            string fullPathToXll = Path.Combine(fullPathInstallFolder, szXllToRegister);

                            RegistryKey rkExcelXll = Registry.CurrentUser.OpenSubKey(szKeyName, true);
                            if (rkExcelXll != null)
                            {
                                Console.WriteLine("Success finding HKCU key for : " + szKeyName);
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
                                        rkExcelXll.SetValue("OPEN" + (nMaxOpen + 1).ToString(CultureInfo.InvariantCulture), "/R \"" + fullPathToXll + "\"");
                                    }
                                    rkExcelXll.Close();
                                }
                                success = true;
                            }
                            else
                            {
                                Console.WriteLine("Unable to retrieve key for : " + szKeyName);
                            }
                        }
                        else
                        {
                            Console.WriteLine("Unable to retrieve registry Information for : " + SzBaseAddInKey +szOfficeVersionKey);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("The list of supported office version is empty");
                }

                Console.WriteLine("End CreateOpenHKCUKey");
            }
            catch (System.Security.SecurityException ex)
            {
                Console.WriteLine("CreateOpenHKCUKey SecurityException" + ex.Message);
                success = false;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine("CreateOpenHKCUKey UnauthorizedAccessException" + ex.Message);
                success = false;
            }
            catch (Exception ex)
            {
                Console.WriteLine("CreateOpenHKCUKey Exception" + ex.Message);
                success = false;
            }
            return success;
        }


        public static bool RemoveHKCUOpenKey(string szOfficeRegKeyVersions, string szXll32Bit, string szXll64Bit)
        {
            bool bFoundOffice = false;
            try
            {
                Console.WriteLine("Begin RemoveHKCUOpenKey");


                if (szOfficeRegKeyVersions.Length > 0)
                {
                    List<string> lstVersions = szOfficeRegKeyVersions.Split(',').ToList();

                    foreach (string szOfficeVersionKey in lstVersions)
                    {
                        // only remove keys where office version is found
                        if (Registry.CurrentUser.OpenSubKey(SzBaseAddInKey + szOfficeVersionKey, false) != null)
                        {
                            bFoundOffice = true;

                            string szKeyName = SzBaseAddInKey + szOfficeVersionKey + @"\Excel\Options";

                            RegistryKey rkAddInKey = Registry.CurrentUser.OpenSubKey(szKeyName, true);
                            if (rkAddInKey != null)
                            {
                                string[] szValueNames = rkAddInKey.GetValueNames();

                                foreach (string szValueName in szValueNames)
                                {
                                    //unregister both 32 and 64 xll
                                    if (szValueName.StartsWith("OPEN") &&
                                        (rkAddInKey.GetValue(szValueName).ToString().Contains(szXll32Bit) ||
                                         rkAddInKey.GetValue(szValueName).ToString().Contains(szXll64Bit)))
                                    {
                                        rkAddInKey.DeleteValue(szValueName);
                                    }
                                }
                            }
                        }
                    }
                }

                Console.WriteLine("End RemoveHKCUOpenKey");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return bFoundOffice;
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
