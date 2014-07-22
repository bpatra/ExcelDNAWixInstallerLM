using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace manageOpenKey
{
    static class HkcuKeys
    {
        public const string SzBaseAddInKey = @"Software\Microsoft\Office\";

        public static void CreateOpenHkcuKey(Parameters parameters)
        {
            if (parameters.SupportedOfficeVersion.Count == 0)
            {
                throw new ApplicationException("There should be at least one office version supported");
            }

            foreach (string szOfficeVersionKey in parameters.SupportedOfficeVersion)
            {
                double nVersion = double.Parse(szOfficeVersionKey, NumberStyles.Any, CultureInfo.InvariantCulture);

                Console.WriteLine("Retrieving Registry Information for : " + SzBaseAddInKey + szOfficeVersionKey);

                // get the OPEN keys from the Software\Microsoft\Office\[Version]\Excel\Options key, skip if office version not found.
                if (Registry.CurrentUser.OpenSubKey(SzBaseAddInKey + szOfficeVersionKey, false) != null)
                {
                    string szKeyName = SzBaseAddInKey + szOfficeVersionKey + @"\Excel\Options";

                    string szXllToRegister = GetAddInName(parameters.XllName, parameters.Xll64Name, szOfficeVersionKey, nVersion);
                    //for a localmachine install the xll's should be in the installFolder
                    string fullPathToXll = Path.Combine(parameters.InstallDirectory, szXllToRegister);

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
                    }
                    else
                    {
                        Console.WriteLine("Unable to retrieve key for : " + szKeyName);
                    }
                }
                else
                {
                    Console.WriteLine("Unable to retrieve registry Information for : " + SzBaseAddInKey + szOfficeVersionKey);
                }
            }

            Console.WriteLine("End CreateOpenHKCUKey");
        }


        public static void RemoveHkcuOpenKey(Parameters parameters)
        {
            Console.WriteLine("Begin RemoveHKCUOpenKey");

            if (parameters.SupportedOfficeVersion.Count == 0)
            {
                throw new ApplicationException("There should be at least one office version supported");
            }
                
            foreach (string szOfficeVersionKey in parameters.SupportedOfficeVersion)
            {
                // only remove keys where office version is found
                if (Registry.CurrentUser.OpenSubKey(SzBaseAddInKey + szOfficeVersionKey, false) != null)
                {

                    string szKeyName = SzBaseAddInKey + szOfficeVersionKey + @"\Excel\Options";

                    using (RegistryKey rkAddInKey = Registry.CurrentUser.OpenSubKey(szKeyName, true))
                    {
                        if (rkAddInKey != null)
                        {
                            string[] szValueNames = rkAddInKey.GetValueNames();

                            foreach (string szValueName in szValueNames)
                            {
                                //unregister both 32 and 64 xll
                                if (szValueName.StartsWith("OPEN") &&
                                    (rkAddInKey.GetValue(szValueName).ToString().Contains(parameters.Xll64Name) ||
                                        rkAddInKey.GetValue(szValueName).ToString().Contains(parameters.XllName)))
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

        //Using a registry key of outlook to determine the bitness of office may look like weird but that's the reality.
        //http://stackoverflow.com/questions/2203980/detect-whether-office-2010-is-32bit-or-64bit-via-the-registry
        private static string GetAddInName(string szXll32Name, string szXll64Name, string szOfficeVersionKey, double nVersion)
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
            else //excel 2003, there was no 64bits
            {
                szXllToRegister = szXll32Name;
            }

            return szXllToRegister;
        }
    }
}
