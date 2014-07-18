using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;

namespace mySampleWiXSetupCA
{
    public class CustomActions
    {

        //private static readonly UIntPtr RegistryArea = Reg.CurrentUser;
        
        private const string Office2007VersionKey = @"Software\Microsoft\Office\12.0";
        private const string Office2010VersionKey = @"Software\Microsoft\Office\14.0";
        private const string Office2013VersionKey = @"Software\Microsoft\Office\15.0";
        public static List<string> OfficeVersionsList = new List<string> { Office2007VersionKey, Office2010VersionKey, Office2013VersionKey };

        [CustomAction]
        public static ActionResult CaUnRegisterAddIn(Session session)
        {
            throw new NotImplementedException();
        }


        //private static ActionResult GetOfficeBitness(Session session, string officeVersionKey, out Reg.BitnessType bitness)
        //{
            //bitness = Reg.BitnessType.Wow6432;//By default and always true for Office2007
            //if (officeVersionKey != Office2007VersionKey)
            //{
            //    StringBuilder myBitness;
            //    var result = Reg.GetRegistryKeyValue(Reg.LocalMachine, Reg.BitnessType.NotApplicable, officeVersionKey + @"\Outlook\", "Bitness", out myBitness);
            //    session.Log(string.Format("1: Reg.LocalMachine, Reg.BitnessType.NotApplicable, officeVersionKey={0}, bitness={1}, result={2}", officeVersionKey, bitness, result));

            //    //Not found, lets try in the 32 bit area
            //    if (result == Reg.ErrorFileNotFound)
            //    {
            //        result = Reg.GetRegistryKeyValue(Reg.LocalMachine, Reg.BitnessType.Wow6432, officeVersionKey + @"\Outlook\", "Bitness", out myBitness);
            //        session.Log(string.Format("2: Reg.LocalMachine, Reg.BitnessType.Wow6432, officeVersionKey={0}, bitness={1}, result={2}", officeVersionKey, bitness, result));

            //        //Not found, lets try in the 64 bit area
            //        if (result == Reg.ErrorFileNotFound)
            //        {
            //            result = Reg.GetRegistryKeyValue(Reg.LocalMachine, Reg.BitnessType.Wow6464, officeVersionKey + @"\Outlook\", "Bitness", out myBitness);
            //            session.Log(string.Format("3: Reg.LocalMachine, Reg.BitnessType.Wow6432, officeVersionKey={0}, bitness={1}, result={2}", officeVersionKey, bitness, result));

            //            //Bitness not found, we initialize to x86
            //            if (result != Reg.ErrorSuccess)
            //            {
            //                myBitness = new StringBuilder("x86");
            //                result = Reg.ErrorSuccess;
            //                session.Log("Bitness could not be found in LocalMachine registry, fallback to x86 bitness");
            //            }
            //        }
            //    }

            //    if (result != Reg.ErrorSuccess)
            //        return ActionResult.Failure;
            //    if (myBitness.ToString() == "x64")
            //    {
            //        bitness = RegistryHelpers.BitnessType.Wow6464;
            //        session.Log("Bitness forced to BitnessType.Wow6464");
            //    }
            //}
            //return ActionResult.Success;
        //}
    }
}
