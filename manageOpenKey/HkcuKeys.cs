﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace manageOpenKey
{
    class HkcuKeys
    {
        private readonly IWriter _writer;
        public HkcuKeys(IWriter writer)
        {
            _writer = writer;
        }

        public const string SzBaseAddInKey = @"Software\Microsoft\Office\";

        public void CreateOpenHkcuKey(Parameters parameters)
        {
            if (parameters.SupportedOfficeVersion.Count == 0)
            {
                throw new ApplicationException("There should be at least one office version supported");
            }
            var registryAdapator = new RegistryAbstractor(_writer);

            foreach (string szOfficeVersionKey in parameters.SupportedOfficeVersion)
            {
                double nVersion = double.Parse(szOfficeVersionKey, NumberStyles.Any, CultureInfo.InvariantCulture);

                _writer.WriteLine("Retrieving Registry Information for : " + SzBaseAddInKey + szOfficeVersionKey);

                // get the OPEN keys from the Software\Microsoft\Office\[Version]\Excel\Options key, skip if office version not found.
                string excelBaseKey = SzBaseAddInKey + szOfficeVersionKey + "Excel";
                    //Software\Microsoft\Office\[Version]\Excel
                if (Registry.LocalMachine.OpenSubKey(excelBaseKey, false) != null)//this version is install on the Machine
                {
                    string excelOptionKey = SzBaseAddInKey + szOfficeVersionKey + @"\Options";

                    //It is very important to Open or Create see https://github.com/bpatra/ExcelDNAWixInstallerLM/issues/9
                    using (RegistryKey rkExcelXll = registryAdapator.OpenOrCreateHkcuKey(excelOptionKey))
                    {
                        string szXllToRegister = GetAddInName(parameters.XllName, parameters.Xll64Name, szOfficeVersionKey,
                                                        nVersion);
                        //for a localmachine install the xll's should be in the installFolder
                        string fullPathToXll = Path.Combine(parameters.InstallDirectory, szXllToRegister);



                        _writer.WriteLine("Success finding HKCU key for : " + excelOptionKey);
                        string[] szValueNames = rkExcelXll.GetValueNames();
                        bool bIsOpen = false;
                        int nMaxOpen = -1;

                        // check every value for OPEN keys
                        foreach (string szValueName in szValueNames)
                        {
                            _writer.WriteLine(string.Format("Examining value {0}", szValueName));
                            // if there are already OPEN keys, determine if our key is installed
                            if (szValueName.StartsWith("OPEN"))
                            {
                                int nOpenVersion = int.TryParse(szValueName.Substring(4), out nOpenVersion) ? nOpenVersion : 0;
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
                                    _writer.WriteLine("Already found the OPEN key " + excelOptionKey);
                                }
                            }
                        }


                        // if adding a new key
                        if (!bIsOpen)
                        {
                            string value = "/R \"" + fullPathToXll + "\"";
                            string keyToUse;
                            if (nMaxOpen == -1)
                            {
                                keyToUse = "OPEN";
                            }
                            else
                            {
                                keyToUse = "OPEN" + (nMaxOpen + 1).ToString(CultureInfo.InvariantCulture);

                            }
                            rkExcelXll.SetValue(keyToUse, value);
                            _writer.WriteLine("Set {0} key with {1} value", keyToUse, value);
                        }
                    }
                }
                else
                {
                    _writer.WriteLine("Unable to retrieve HKLM key for: " + excelBaseKey);
                }
            }

            _writer.WriteLine("End CreateOpenHKCUKey");
        }


        public void RemoveHkcuOpenKey(Parameters parameters)
        {
            _writer.WriteLine("Begin RemoveHKCUOpenKey");

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


            _writer.WriteLine("End RemoveHKCUOpenKey");
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
