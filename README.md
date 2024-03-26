# SetPEImageProps
Set PE Image Header Properties

modSetPEImageProps.twin is a module designed to be dropped into a twinBASIC project, after which it automatically modifies the compiled binaries right after they're built, using tB's `[RunAfterBuild]` attribute. 

I wanted to investigate whether the default version values tB applies were causing issues, so made this as a temporary solution until these options can be set by tB itself. Currently this module can adjust the following in the `IMAGE_OPTIONAL_HEADER`:

* Linker version
* Operating system version
* Image version
* Subsystem version
* Win32VersionValue (should probably not change)
* DllCharacteristics

  twinBASIC sets the linker version to 6.0, which is associated with Visual Studio 6.0, and what I think might be more of an issue, it sets the OS version and subsystem version to 4.0 (Windows NT 4.0), while these are almost always set higher when compiling with other tools. Image version is often set to the same as those two, but is sometimes left as 1.0 (tB default as well). With the configuration provided in the code as posted here, the linker is set to 7.0 just to disassociate it from VS6, and the OS/Image/Subsystem versions are all set to 10.0. tB already supports setting the actual subsystem (GUI, CUI, or Native), so that's not currently included.

  The other set of options included is to modify the `DllCharacteristics` field. I need to set one of these for one of my driver projects, but usually you should just leave it alone. Just adding a flag doesn't add support for a feature, so if you e.g. enable the control flow guard, it will likely result in crashing.

  The design is fairly straightforward, should you want to modify any other fields. Again, this just hard codes values in the header, it won't change the code, and you don't want a mismatch. The actual function takes an optional path argument so you can call it manually at any time to adjust any file, but the automatic modification was my main goal here. It can handle bitness mismatches if you do call it manually.

![image](https://github.com/fafalone/SetPEImageProps/assets/7834493/77e66fd9-9d18-4e89-80aa-47770129ba4d)

**Requirements**
This module relies on my Windows Development Library for twinBASIC package being added.

  > [!IMPORTANT]
  > **Requires an updated WinDevLib!** If you drop the module into an existing project, you must also update WinDevLib if you're already using it, as there were bugs that would result in a corrupt exe using this in 64bit mode.

---

Configuration:

  ```vba
  #Const MASTER_ENABLE = 1

#Const ENABLE_OPT_HEADER_VERSIONING = 1
Private Const MajorLinkerVersion As Byte = 7 'Default: 6
Private Const MinorLinkerVersion As Byte = 0 'Default: 0
Private Const MajorOperatingSystemVersion As Integer = 10 'Default: 4 (4.0 - Windows NT 4.0)
Private Const MinorOperatingSystemVersion As Integer = 0 'Default: 0
Private Const MajorImageVersion As Integer = 10 'Default: 1
Private Const MinorImageVersion As Integer = 0 'Default: 0
Private Const MajorSubsystemVersion As Integer = 10 'Default: 4 (4.0 - Windows NT 4.0)
Private Const MinorSubsystemVersion As Integer = 0 'Default: 0
Private Const Win32VersionValue As Long = 0 'Default: 0

#Const ENABLE_OPT_HEADER_DLL_CHARACTERISTICS = 0
Private Const DLLCHARACTERISTICS_HIGH_ENTROPY_VA = 0 'Default: 0  ' Image can handle a high entropy 64-bit virtual address space.
Private Const DLLCHARACTERISTICS_DYNAMIC_BASE = 1 'Default: 1  ' DLL can move.
Private Const DLLCHARACTERISTICS_FORCE_INTEGRITY = 0 'Default: 0  ' Code Integrity Image
Private Const DLLCHARACTERISTICS_NX_COMPAT = 0 'Default: 0  ' Image is NX compatible
Private Const DLLCHARACTERISTICS_NO_ISOLATION = 0 'Default: 0  ' Image understands isolation and doesn't want it
Private Const DLLCHARACTERISTICS_NO_SEH = 1 'Default: 1 ' Image does not use SEH.  No SE handler may reside in this image
Private Const DLLCHARACTERISTICS_NO_BIND = 0 'Default: 0  ' Do not bind this image.
Private Const DLLCHARACTERISTICS_APPCONTAINER = 0 'Default: 0  ' Image should execute in an AppContainer
Private Const DLLCHARACTERISTICS_WDM_DRIVER = 0 'Default: 0  ' Driver uses WDM model
Private Const DLLCHARACTERISTICS_GUARD_CF = 0 'Default: 0  ' Image supports Control Flow Guard.
Private Const DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE = 0 'Default: 0
```

