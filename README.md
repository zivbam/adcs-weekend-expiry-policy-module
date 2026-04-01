# ADCS Midweek Expiry Policy Module

A custom **Active Directory Certificate Services (ADCS)** policy module that enforces certificate expiration (`NotAfter`) on **Monday** or **Tuesday**.

If ADCS computes an expiration on another day, the module shortens validity and moves expiration to the **previous Tuesday**.

## Rule implemented

- Monday: keep as-is.
- Tuesday: keep as-is.
- Wednesday/Thursday/Friday/Saturday/Sunday: move back to previous Tuesday.

This behavior is deterministic and easy to audit.

## What is included in this repository

- `adcs-weekend-expiry-policy-module.sln`
- `MidweekExpiryPolicy/MidweekExpiryPolicy.vcxproj`
- COM policy module implementation (`ICertPolicy2`) in C++.
- COM registration exports (`DllRegisterServer`, `DllUnregisterServer`).

## Module identity

- **CLSID:** `{B4AB7B5E-BD14-4F36-93B9-F4A5F1B6EA2B}`
- **ProgID:** `MidweekExpiryPolicy.Module`
- **Output DLL name:** `MidweekExpiryPolicy.dll`

## Prerequisites

- Windows Server with ADCS installed.
- Visual Studio 2022 with C++ workload.
- Windows SDK that includes ADCS headers/libs.
- Lab environment for validation.

## Build

### Visual Studio

1. Open `adcs-weekend-expiry-policy-module.sln`.
2. Select `Release | x64`.
3. Build Solution.

Expected artifact:

```text
x64\Release\MidweekExpiryPolicy.dll
```

### MSBuild

```powershell
msbuild .\adcs-weekend-expiry-policy-module.sln /t:Build /p:Configuration=Release /p:Platform=x64
```

## Deploy to CA server

> Test in lab first and back up CA configuration before changing policy modules.

1. Copy DLL:

```powershell
Copy-Item .\x64\Release\MidweekExpiryPolicy.dll C:\Windows\System32\CertSrv\ -Force
```

2. Register COM server:

```powershell
regsvr32 C:\Windows\System32\CertSrv\MidweekExpiryPolicy.dll
```

3. Configure CA policy module CLSID:

```powershell
certutil -setreg policy\PolicyModules "{B4AB7B5E-BD14-4F36-93B9-F4A5F1B6EA2B}"
```

4. Restart ADCS:

```powershell
Restart-Service CertSvc
```

## Validate behavior

1. Submit test certificate requests from relevant templates.
2. Inspect `NotAfter` day of week.
3. Confirm day is only Monday or Tuesday.

Useful command:

```powershell
certutil -dump .\issued-test.cer
```

## Rollback

1. Restore previous policy module registry value.
2. Restart `CertSvc`.
3. Unregister this module if needed:

```powershell
regsvr32 /u C:\Windows\System32\CertSrv\MidweekExpiryPolicy.dll
```

## Notes

- Effective lifetime can be shorter than template nominal validity.
- The current implementation applies the rule during `VerifyRequest` by adjusting certificate property `ExpirationDate`.
- Keep identical version and policy behavior across all issuing CAs.
