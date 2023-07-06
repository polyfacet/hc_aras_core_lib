# hc_aras_core_lib

Aras Core Extension

Generic server extension functions for Aras Innovator

This has been upgraded to .NET Core 6 from previous .NET Framework (For .NET Framework versions, use the `Last-Net-Framework-Version`)

## To build

- Requires dotnet 6.0 sdk
- Your own explicit references to the corresponding
  - ..\..\References\ArasR2023\IOM.dll
  - ..\..\References\ArasR2023\Aras.Server.Core.dll

## To run BOM Compare (xUnit) tests

- As for the build you need to have your own "SDK"-IOM.dll reference.  
  This can be picked from the Export or Import tool.
