ExcelDNAWixInstallerLM
======================

## A sample Wix installer "per machine" for Excel-DNA addins.

The [Excel-DNA Wix installer][wixinstallerurl] sample found is per User which makes it not suitable for install at "company level". Then we have to perform a per Machine install. This cannot be done directly with HKLM because the OPEN key for .xll can be only in the HKCU registry. Therefore this install uses the ActiveSetup feature of Windows that propagates some install for all users of a given machine. The user who wants technical details can check [my blogpost on the subject][blogpost]

This project may also be interesting for people looking for a Wix installer that uses the ActiveSetup feature (there are some traps ahead...).

## What you should know before reusing this sample for your own project

###Configuration
Then only configuration that you are supposed to do is on the *.wxi* file. Check the instruction there.
You MUST change the GUIDs there and replace by yours. To generate a GUID you may type [System.Guid]::NewGuid() in a Powershell Console.

###Requirements for this sample
In this sample we have set only 12.0,14.0,15.0 (Office 2007, 2010 and 2014). But there should be a big deal to support 2003 even we have not tested yet.
In this install the Full .Net Framework 4.0 is required. However we believe that you can require .NET 4.0 Client Profile because many elements of the Microsoft.Win32 namespace used in this sample are present in .NET40 Client Profile.
Support for .NET 3.5 will require to rewrite a good proportion of the .NET Registry APIs.
 
###A summary of the solution involved there.
The Wix installer has only Install, Repair, Uninstall (no Modify). It modifies the HKLM registry then it requires elevated (Administrator) privileges.
While installing and repairing, the msi installs the two xlls (one for Office32 and one for Office64) in the x86 ProgamFiles. Then for the current user invokes an executable (called *manageOpenKey.exe*) that will set the proper OPEN HKCU key for this current user. Therefore, there is no need for the current user to reboot for using the addin. The other user install is handled using the ActiveSetup. The ActiveSetup is an HKLM registry key that is "mirorred" in the HKCUs. If not present in the HKCU, which will happen when a new user logs in, then a script of your choice can be executed. Here we invoke the same *manageOpenKey.exe* mentioned before. Remark that, the upgrading (versioning) and uninstall is also handled [see more here][blogpost]

###Limitations
There is only one limitation known for now in the uninstall. The uninstall does not wipe completly the *[INSTALLFOLDER]*. Indeed, we have to leave the *manageOpenKey.exe* which is also responsible for cleaning the environmment for *all* users at uninstall.

After uninstall, for other users than the admin user who has performed the uninstall, a pop-up is launched asking a message of the form *"<Your product>" by "<Your Company>" has been removed from this computer. Do you want to clean up you personalized settings for this program?* (see the popup, remark that the *This is an example* will be replaced by  *"<Your product>" by "<Your Company>"*)
. User, will have to click Yes (if not it will ask at each logon).
![alt text](https://cloud.githubusercontent.com/assets/2801702/3671277/f998e5ba-124d-11e4-9382-c38e869401e7.jpg "Pop up Uninstalled ActiveSetup")

###Others
One suprisingly fact: if a non admin tries to uninstall, then the OS will ask to execute throught administrator privileges (then another user). If he does so then, the addin will not be properly uninstalled for the non admin user. He will have to relog so that the uninstall process per user triggered by ActiveSetup is executed (and the OPEN key will be removed).

[wixinstallerurl]: https://github.com/Excel-DNA/WiXInstaller "ExcelDNA Wix installer"
[blogpost]: http://www.notwrittenyet.com "TODO: update"

