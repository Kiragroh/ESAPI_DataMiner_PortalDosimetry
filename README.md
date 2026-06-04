# DataMiner_PortalDosimetry
## General ESAPI Guidance

For current build/runtime notes, clinical-safety reminders, stand-alone threading rules, write-enabled script governance, and ESAPI 18.x considerations, see [`docs/ESAPI_GENERAL_NOTES.md`](docs/ESAPI_GENERAL_NOTES.md).

(Standalone-Console-App)
Mine all analysed PortalDosimetry Fields and make custom evaluations too. Have Fun.

Installation notes:

1.) Add three .dll-files to your dependecies normally found in C:\Program Files (x86)\Varian\ProductLine\Workspaces\VMS.PortalDosimetry.Workspace\
- PortalDosimetry.dll
- VMS.CA.Scripting.dll
- VMS.DV.PD.Scripting.dll

2.) Change Network-Path in Line 310 and 312 to your liking. Otherwise your file will be created in Temp-Path

Note:
- script is optimized to work with Eclipse 15.1
- absolute beginner should first read my beginnerGuide
https://drive.google.com/drive/folders/1-aYUOIfyvAUKtBg9TgEETiz4SYPonDOO
