# Change Log

[v1.3.0.0, 11 Aug 2026]

 - Fixed CJK/Japanese/Korean/emoji filenames being wrongly rejected as invalid (signed AscW comparison)
 - Reserved device names (CON, PRN, NUL, COM1-9, LPT1-9) now rejected uniformly as bad names (error 52) on all name-taking operations
 - Path comparisons now use a locale-independent case fold, fixing potential MergeFiles/CopyFolder misbehavior under Turkish/Azeri locales
 - Consolidated name validation into a single FSTShared.GuardLeafName; colon permitted only on the CreateTextFile ADS path
 - Added simple test procedure to GitHub repo to demo functionality and to help faciliate issue resolution

[v1.2.0.0, 09 Aug 2026]

 - Added IsFileLocked to FileSystemTools and FileSystemObject, and IsLocked to File
 - Dropped useLongPathPrefix optional argument for GetAbsolutePathName - handled internally
 
[v1.1.0.1, 07 Aug 2026]

 - Initial Public release
