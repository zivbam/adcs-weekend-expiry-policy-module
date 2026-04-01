# ADCS Midweek Expiry Policy Module

A custom **Active Directory Certificate Services (ADCS)** policy module that prevents certificate expirations from landing near the weekend by enforcing `NotAfter` on **Monday** or **Tuesday** only.

## Why this exists

In many PKI environments, certificates that expire on Friday/Saturday/Sunday can create operational risk:

- fewer support engineers are available,
- escalation paths are slower,
- business impact can be larger if renewals fail.

This module reduces that risk by shortening validity (when needed) so expiration lands in a safer weekday window.

---

## Behavior

During issuance, ADCS computes validity from template and CA settings. The policy module then evaluates the computed `NotAfter` and adjusts it backward if it does not fall on Monday or Tuesday.

### Default decision rule (recommended)

- If `NotAfter` is **Monday**: keep as-is.
- If `NotAfter` is **Tuesday**: keep as-is.
- If `NotAfter` is **Wednesday, Thursday, Friday, Saturday, or Sunday**: move back to the **previous Tuesday**.

This provides a deterministic and easy-to-explain rule.


### Important consequence

The resulting certificate lifetime may be shorter than the template default by up to several days.

---

## Technical architecture

The module is implemented as a COM policy DLL for ADCS, typically using:

- `ICertPolicy` (legacy)
- `ICertPolicy2` (recommended)

The key logic is in `VerifyRequest`, where the module:

1. Reads request/context validity values,
2. Calculates adjusted `NotAfter`,
3. Writes back effective validity,
4. Returns issue/deny/defer decision.

---

## Prerequisites

- Windows Server with **ADCS** role installed.
- Administrative access to the CA server.
- Visual Studio 2022 (recommended) with:
  - Desktop development with C++
  - Compatible Windows SDK
- A non-production lab CA for validation before rollout.

---

## Build instructions

These steps assume a C++ DLL project.

### Build in Visual Studio

1. Open solution file:
   - `adcs-weekend-expiry-policy-module.sln`
2. Set configuration:
   - `Release`
   - `x64`
3. Run **Build Solution**.

### Build via command line (MSBuild)

```powershell
msbuild .\adcs-weekend-expiry-policy-module.sln /t:Build /p:Configuration=Release /p:Platform=x64
```

Expected output (example):

```text
.\x64\Release\MidweekExpiryPolicy.dll
```

---

## Deployment to ADCS

> Always test in lab first. Back up CA configuration before changes.

### 1) Copy the module DLL

```powershell
Copy-Item .\x64\Release\MidweekExpiryPolicy.dll C:\Windows\System32\CertSrv\ -Force
```

### 2) Register COM server

```powershell
regsvr32 C:\Windows\System32\CertSrv\MidweekExpiryPolicy.dll
```

### 3) Configure CA to use the policy module

Set policy module registration values according to your module's CLSID/ProgID registration.

Example (illustrative):

```powershell
certutil -setreg policy\PolicyModules "{YOUR-POLICY-MODULE-CLSID}"
```

> Exact registry path/value depends on how your module registers itself (`DllRegisterServer` behavior).

### 4) Restart ADCS service

```powershell
Restart-Service CertSvc
```

---

## Validation checklist

After deployment:

1. Submit test requests from one or more templates.
2. Inspect issued certificate `NotAfter`.
3. Confirm expiration day is Monday or Tuesday only.

Useful command:

```powershell
certutil -dump .\issued-test.cer
```


## Production considerations

- **Shorter effective validity:** expected behavior, not a bug.
- **Compliance alignment:** verify CP/CPS and internal policy allow shortened validity.
- **Observability:** log each adjustment (original vs adjusted `NotAfter`) to Event Log.
- **Operational safety:** support feature flag or registry switch for emergency disablement.
- **Consistency:** keep identical module version/rules across all issuing CAs.

---

## Rollback procedure

If issues occur:

1. Restore previous policy module configuration.
2. Restart `CertSvc`.
3. Optionally unregister new DLL:
   - `regsvr32 /u <path-to-dll>`
4. Restore previous DLL from backup.

---

## Troubleshooting

- `regsvr32` fails:
  - verify x64/x86 alignment,
  - verify runtime dependencies.
- CA service fails after policy change:
  - revert policy registry values,
  - restart `CertSvc`.
- No `NotAfter` change observed:
  - verify module is loaded,
  - verify `VerifyRequest` path is executed,
  - inspect Event Log/debug logging.
