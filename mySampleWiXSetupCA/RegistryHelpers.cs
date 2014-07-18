//This helper is designed to access the registry with .NET 3.5
//It calls the low level native Win32 API and is therefore able to access to WOW64_32KEY and WOW64_64KEY sections

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace mySampleWiXSetupCA
{
    public class RegistryHelpers
    {
        public enum BitnessType { NotApplicable = 0, Wow6432 = 0x0200, Wow6464 = 0x0100 }
        public enum ValueType : uint { RegNone = 0, RegString = 1, RegExpandSting = 2, RegBinary = 3, RegDword = 4, RegDwordLittleEndian = 4, RegDwordBigEndian = 5, RegLink = 6, RegMultiString = 7 }
        public static readonly UIntPtr LocalMachine = new UIntPtr(0x80000002u);
        public static readonly UIntPtr CurrentUser = new UIntPtr(0x80000001u);
        public const int ErrorSuccess = 0x0;
        public const int ErrorFileNotFound = 0x2;

        /// <summary>
        /// Determines if the key exist or no.
        /// </summary>
        /// <param name="registryArea">LocalMachine or CurrentUser</param>
        /// <param name="bitnessType">NotApplicable, Wow6432 or Wow6464</param>
        /// <param name="keyPath">Exact registry path to the key</param>
        /// <returns>0 if ok, error othewise</returns>
        public static int ExistRegistryKey(UIntPtr registryArea, BitnessType bitnessType, string keyPath)
        {
            var hKeyVal = UIntPtr.Zero;
            try
            {
                //Open key for reading
                return RegOpenKeyEx(registryArea, keyPath, 0, KeyQueryValue | (int)bitnessType, out hKeyVal);
            }
            finally
            {
                if (hKeyVal != UIntPtr.Zero)
                {
                    RegCloseKey(hKeyVal);
                }
            }

        }

        /// <summary>
        /// Write Data to the registry under the Name. The Key must exist before.
        /// </summary>
        /// <param name="registryArea">LocalMachine or CurrentUser</param>
        /// <param name="bitnessType">NotApplicable, Wow6432 or Wow6464</param>
        /// <param name="keyPath">Exact registry path to the key</param>
        /// <param name="valueName">Value Name</param>
        /// <param name="valueType">Value Type </param>
        /// <param name="value">Value to write</param>
        /// <returns>0 if ok, error othewise</returns>
        public static int WriteRegistryKeyValue(UIntPtr registryArea, BitnessType bitnessType, string keyPath, string valueName, ValueType valueType, StringBuilder value)
        {
            var hKeyVal = UIntPtr.Zero;
            try
            {
                //Open key for writing
                var result = RegOpenKeyEx(registryArea, keyPath, 0, KeySetValue | (int)bitnessType, out hKeyVal);
                if (result == 0)
                {
                    result = RegSetValueEx(hKeyVal, valueName, 0, (uint)valueType, value, value.Length + 1);
                }
                return result;
            }
            finally
            {
                if (hKeyVal != UIntPtr.Zero)
                {
                    RegCloseKey(hKeyVal);
                }
            }
        }

        /// <summary>
        /// Create a key in the Registry
        /// </summary>
        /// <param name="registryArea">LocalMachine or CurrentUser</param>
        /// <param name="bitnessType">NotApplicable, Wow6432 or Wow6464</param>
        /// <param name="keyPath">Exact registry path to the key</param>
        /// <returns>0 if ok, error othewise</returns>
        public static int CreateRegistryKey(UIntPtr registryArea, BitnessType bitnessType, string keyPath)
        {
            var hKeyVal = UIntPtr.Zero;
            try
            {
                RegistryDispositionValue regDisValue;
                var result = RegCreateKeyEx(registryArea, keyPath, 0, null, 0, KeyAllAccess | (int)bitnessType, UIntPtr.Zero, out hKeyVal, out regDisValue);
                return result;
            }
            finally
            {
                if (hKeyVal != UIntPtr.Zero)
                {
                    RegCloseKey(hKeyVal);
                }
            }
        }

        /// <summary>
        /// Delete a key from the Registry
        /// </summary>
        /// <param name="registryArea">LocalMachine or CurrentUser</param>
        /// <param name="bitnessType">NotApplicable, Wow6432 or Wow6464</param>
        /// <param name="keyPath">Exact registry path to the key</param>
        /// <returns>0 if ok, error othewise</returns>
        public static int DeleteRegistryKey(UIntPtr registryArea, BitnessType bitnessType, string keyPath)
        {
            try
            {
                var result = RegDeleteKeyEx(registryArea, keyPath, (int)bitnessType, 0);
                return result;
            }
            catch
            {
                throw new Exception("Fatal error deleting registry key '" + keyPath + "'");
            }
        }

        /// <summary>
        /// Delete a key from the Registry
        /// </summary>
        /// <param name="registryArea">LocalMachine or CurrentUser</param>
        /// <param name="bitnessType">NotApplicable, Wow6432 or Wow6464</param>
        /// <param name="keyPath">Exact registry path to the key</param>
        /// <param name="valueName">Value Name to delete</param>
        /// <returns>0 if ok, error othewise</returns>
        public static int DeleteRegistryKeyValue(UIntPtr registryArea, BitnessType bitnessType, string keyPath, string valueName)
        {
            var hKeyVal = UIntPtr.Zero;
            try
            {
                //Open key for writing
                var result = RegOpenKeyEx(registryArea, keyPath, 0, KeySetValue | (int)bitnessType, out hKeyVal);
                if (result == 0)
                {
                    result = RegDeleteKeyValue(hKeyVal, "", valueName);
                }
                return result;
            }
            finally
            {
                if (hKeyVal != UIntPtr.Zero)
                {
                    RegCloseKey(hKeyVal);
                }
            }
        }

        /// <summary>
        /// Lists all the Subkeys of a Key
        /// </summary>
        /// <param name="registryArea">LocalMachine or CurrentUser</param>
        /// <param name="bitnessType">NotApplicable, Wow6432 or Wow6464</param>
        /// <param name="keyPath">Exact registry path to the key</param>
        /// <param name="subKeys">returns list of all the Subkeys</param>
        /// <returns>0 if ok, error othewise</returns>
        public static int GetRegistryKeySubKeys(UIntPtr registryArea, BitnessType bitnessType, string keyPath, out List<string> subKeys)
        {
            var hKeyVal = UIntPtr.Zero;
            subKeys = new List<string>();
            try
            {
                //Open key 
                var result = RegOpenKeyEx(registryArea, keyPath, 0, KeyQueryValue | (int)bitnessType, out hKeyVal);
                if (result == 0)
                {
                    var i = 0;
                    while (result == 0 || result != ErrorNoMoreItems)
                    {
                        var keyLength = MaxRegKeynameSize + 1;
                        long lastTimeWrite;
                        var subKey = new StringBuilder(0, MaxRegKeynameSize + 1);
                        result = RegEnumKeyEx(hKeyVal, i, subKey, ref keyLength, 0, IntPtr.Zero, IntPtr.Zero, out lastTimeWrite);
                        if (result == 0 || result != ErrorNoMoreItems)
                        {
                            subKeys.Add(subKey.ToString());
                            i++;
                        }
                    }
                }
                if (result == ErrorNoMoreItems)
                    result = 0;
                return result;
            }
            finally
            {
                if (hKeyVal != UIntPtr.Zero)
                {
                    RegCloseKey(hKeyVal);
                }
            }
        }

        /// <summary>
        /// Retrieve a Value of a Key
        /// </summary>
        /// <param name="registryArea">LocalMachine or CurrentUser</param>
        /// <param name="bitnessType">NotApplicable, Wow6432 or Wow6464</param>
        /// <param name="keyPath">Exact registry path to the key</param>
        /// <param name="valueName">Value name to retrieve</param>
        /// <param name="value">Corresponding value</param>
        /// <returns>0 if ok, error othewise</returns>
        public static int GetRegistryKeyValue(UIntPtr registryArea, BitnessType bitnessType, string keyPath, string valueName, out StringBuilder value)
        {
            var hKeyVal = UIntPtr.Zero;
            try
            {
                //Open key
                var result = RegOpenKeyEx(registryArea, keyPath, 0, KeyQueryValue | (int)bitnessType, out hKeyVal);
                value = new StringBuilder(MaxRegKeynameSize);
                if (result == 0)
                {
                    uint valueType;
                    var lpcbData = value.Capacity;
                    result = RegQueryValueEx(hKeyVal, valueName, 0, out valueType, value, ref lpcbData);
                }
                return result;
            }
            finally
            {
                if (hKeyVal != UIntPtr.Zero)
                {
                    RegCloseKey(hKeyVal);
                }
            }
        }

        /// <summary>
        /// Lists all the Values of a Key
        /// </summary>
        /// <param name="registryArea">LocalMachine or CurrentUser</param>
        /// <param name="bitnessType">NotApplicable, Wow6432 or Wow6464</param>
        /// <param name="keyPath">Exact registry path to the key</param>
        /// <param name="valueNames">Return list of all the value Names</param>
        /// <param name="values">Return list of corresponding values</param>
        /// <returns>0 if ok, error othewise</returns>
        public static int GetRegistryKeyValues(UIntPtr registryArea, BitnessType bitnessType, string keyPath, out List<string> valueNames, out List<StringBuilder> values)
        {
            var hKeyVal = UIntPtr.Zero;
            valueNames = new List<string>();
            values = new List<StringBuilder>();
            try
            {
                //Open key
                var result = RegOpenKeyEx(registryArea, keyPath, 0, KeyQueryValue | (int)bitnessType, out hKeyVal);
                if (result == 0)
                {
                    var i = 0;
                    while (result == 0 || result != ErrorNoMoreItems)
                    {
                        var valueNameLength = MaxRegKeynameSize + 1;
                        var valueLength = MaxRegValueSize + 1;
                        var valueName = new StringBuilder(MaxRegKeynameSize + 1);
                        var value = new StringBuilder(MaxRegValueSize + 1);
                        result = RegEnumValue(hKeyVal, i, valueName, ref valueNameLength, 0, IntPtr.Zero, value, ref valueLength);
                        //Console.WriteLine("After=" + valueName + "(" + valueNameLength + ") Value=" + value + " (" + valueLength + ")");
                        if (result == 0 || result != ErrorNoMoreItems)
                        {
                            valueNames.Add(valueName.ToString());
                            values.Add(value);
                            i++;
                        }
                    }
                }
                if (result == ErrorNoMoreItems)
                    result = 0;
                return result;
            }
            finally
            {
                if (hKeyVal != UIntPtr.Zero)
                {
                    RegCloseKey(hKeyVal);
                }
            }
        }

        private const int KeyQueryValue = 0x0001;
        private const int KeySetValue = 0x0002;
        private const int KeyAllAccess = 0xF003F;
        private const int ErrorNoMoreItems = 259;
        private const int MaxRegKeynameSize = 255;
        private const int MaxRegValueSize = 16383;

        [Flags]
        public enum RegistryDispositionValue : uint
        {
            RegCreatedNewKey = 0x00000001,
            RegOpenedExistingKey = 0x00000002
        }

        [DllImport("advapi32.dll", EntryPoint = "RegOpenKeyEx")]
        private static extern int RegOpenKeyEx(UIntPtr hKey, string lpSubKey, uint ulOptions, int samDesired, out UIntPtr phkResult);

        [DllImport("advapi32.dll", EntryPoint = "RegQueryValueEx")]
        private static extern int RegQueryValueEx(UIntPtr hKey, string lpValueName, int lpReserved, out uint lpType, StringBuilder lpData, ref int lpcbData);

        [DllImport("advapi32.dll", EntryPoint = "RegSetValueEx")]
        private static extern int RegSetValueEx(UIntPtr hKey, string lpValueName, int lpReserved, uint lpType, StringBuilder lpData, int lpcbData);

        [DllImport("advapi32.dll", EntryPoint = "RegCreateKeyEx")]
        private static extern int RegCreateKeyEx(UIntPtr hKey, string lpSubKey, int lpReserved, string lpClass, uint dwOptions, int samDesired, UIntPtr lpSecurityAttributes,
                                                out UIntPtr phkResult, out RegistryDispositionValue lpdwDisposition);

        [DllImport("advapi32.dll", EntryPoint = "RegDeleteKeyEx")]
        private static extern int RegDeleteKeyEx(UIntPtr hKey, string lpSubKey, int samDesired, uint reserved);

        [DllImport("advapi32.dll", EntryPoint = "RegDeleteKeyValue")]
        private static extern int RegDeleteKeyValue(UIntPtr hKey, string lpSubKey, string lpValueName);

        [DllImport("advapi32.dll", EntryPoint = "RegEnumKeyEx")]
        private static extern int RegEnumKeyEx(UIntPtr hKey, int index, StringBuilder lpName, ref int lpcbName, int reserved, IntPtr lpClass, IntPtr lpcbClass, out long lpftLastWriteTime);

        [DllImport("advapi32.dll", EntryPoint = "RegEnumValue")]
        private static extern int RegEnumValue(UIntPtr hKey, int index, StringBuilder lpValueName, ref int lpcValueName, int lpReserved, IntPtr lpType, StringBuilder lpData, ref int lpcbData);

        [DllImport("advapi32.dll", EntryPoint = "RegCloseKey")]
        private static extern int RegCloseKey(UIntPtr hKey);


    }
}
