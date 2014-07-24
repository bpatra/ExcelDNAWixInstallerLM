ExcelDNAWixInstallerLM
======================

# A sample Wix installer "per machine" for Excel-DNA addins.

The [Excel-DNA Wix installer][wixinstallerurl] sample found is per User which makes it not suitable for install at "company level". Then we have to perform a per Machine install. This cannot be done directly with HKLM because the OPEN key for .xll can be only in the HKCU registry. Therefore this install uses the ActiveSetup feature of Windows that propagates some install for all users of a given machine. The user who wants technical details can check [my blogpost on the subject]

This project may also be interesting for people looking for a Wix installer that uses the ActiveSetup feature (there are some traps ahead...).

# What you should know before reusing this project

TODO:

[wixinstallerurl]: https://github.com/Excel-DNA/WiXInstaller "ExcelDNA Wix installer"

