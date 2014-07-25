ExcelDNAWixInstallerLM
======================

## A sample Wix installer "per machine" for Excel-DNA addins.

The [Excel-DNA Wix installer][wixinstallerurl] sample found is per User which makes it not suitable for install at "company level". Then we have to perform a per Machine install. This cannot be done directly with HKLM because the OPEN key for .xll can be only in the HKCU registry. Therefore this install uses the ActiveSetup feature of Windows that propagates some install for all users of a given machine. The user who wants technical details can check [my blogpost on the subject][blogpost]

This project may also be interesting for people looking for a Wix installer that uses the ActiveSetup feature (there are some traps ahead...).

## What you should know before reusing this sample for your own project

###Configuration
+ Then only configuration that you are supposed to do is on the *.wxi* file under the */mySampleWiXSetup* directory. Check the instruction there.

+ You MUST change the GUIDs there and replace by yours. To generate a GUID you may type [System.Guid]::NewGuid() in a Powershell Console.

###Requirements for this sample
+ In this sample we have set only the list of supported Office version to 12.0,14.0,15.0 (Office 2007, 2010 and 2014). But there should not be a big deal to support 2003 as well even we have tried.

+ In this install the Full .Net Framework 4.0 is required. However we believe that you can require .NET 4.0 Client Profile because many elements of the Microsoft.Win32 namespace used in this sample are present in .NET40 Client Profile. Support for .NET 3.5 will require to rewrite a good proportion of the .NET Registry APIs.
 
###A summary of the solution involved there.
+ The Wix installer has only Install, Repair, Uninstall (no Modify). It modifies the HKLM registry then it requires elevated (Administrator) privileges.

+ While installing and repairing, the msi installs the two Excel-DNA packed xlls (one for Office32 and one for Office64) in the x86 ProgramFiles directory. Then, for the current user, it invokes an executable (called *manageOpenKey.exe*) that will set the proper *HKCU OPEN key* for this current user. **Therefore, there is no need for the current user to reboot for using the addin**. The setup for the other user is handled using the ActiveSetup windows feature. Shortly, this is an HKLM registry key that is "mirorred" in the HKCUs. If not present in the HKCU, which will happen when a user logs in for the first time, then a script of your choice can be executed. Here we invoke the same *manageOpenKey.exe* mentioned before to setup the *HKCU OPEN key*. Note also that the upgrading (versioning) and uninstall is also handled for more information [see here][blogpost]

###Limitations
+ There is only one limitation known for now and its in the uninstall. The uninstall does not wipe completly the *[INSTALLFOLDER] typically (%SystemDrive%/Program Files(x86)/YourCompany/YourProduct)*. Indeed, we have to leave the *manageOpenKey.exe* which is also responsible for cleaning the environmment for *all* users.

###Others
One fact that may surprise: if a non admin tries to uninstall, then the OS will ask to execute throught administrator privileges (then another user). If he does so then, the addin will not be properly uninstalled for the non admin user. He will have to relog so that the uninstall process per user triggered by ActiveSetup is executed (and the OPEN key will be removed).

[wixinstallerurl]: https://github.com/Excel-DNA/WiXInstaller "ExcelDNA Wix installer"
[blogpost]: http://www.notwrittenyet.com "TODO: update"

