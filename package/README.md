# Graph Connectivity Tests

Microsoft Graph connectivity tests in the user context: free/busy in the same organisation, in both directions of a hybrid organisation and across organisations — plus mailbox services, provisioning and the DNS records of a domain.

Author: Nicolas Fabert.

This folder contains everything needed to run the tool: `Invoke-GraphConnectivityTests.ps1`, the module, the configuration, the report template and the guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements
- Windows 10 / 11 or Windows Server 2016 to 2025.
- PowerShell 7.4 or later.
- Microsoft Edge or Google Chrome for the sign-in window, or device code sign-in.
- HTTPS to `login.microsoftonline.com`, `graph.microsoft.com`, `outlook.office365.com`, and to the on-premises Exchange when one is involved.
- The source mailbox and its password. No app registration with the default client (Microsoft Office); an Exchange administrator for `-AdminChecks`; an app registration with a certificate for the Application context.

## Quick start
```powershell
notepad .\config\GraphConnectivityTests.config.psd1     # the source mailbox and the target mailboxes

.\Invoke-GraphConnectivityTests.ps1 -SourceMailbox alice@contoso.com -TargetMailbox bob@contoso.com              # can Alice see Bob's free/busy?
.\Invoke-GraphConnectivityTests.ps1 -SourceMailbox alice@contoso.com -TargetMailbox onprem.user@contoso.com -AdminChecks   # Microsoft 365 to on-premises
.\Invoke-GraphConnectivityTests.ps1 -SourceMailbox onprem.user@contoso.com -TargetMailbox alice@contoso.com       # on-premises to Microsoft 365
.\Invoke-GraphConnectivityTests.ps1 -TestType Provisioning -SourceMailbox alice@contoso.com -TargetMailbox new.user@contoso.com
.\Invoke-GraphConnectivityTests.ps1 -TestType DomainDns -Domain contoso.com -SourceMailbox alice@contoso.com
.\Invoke-GraphConnectivityTests.ps1 -Gui                                                                         # the same in a window
```

## Content
| Item | Role |
|---|---|
| `Invoke-GraphConnectivityTests.ps1` | Entry script. |
| `GraphConnectivityTests.psd1` | Module manifest. |
| `GraphConnectivityTests.psm1` | Module loader. |
| `config\` | Example configuration. |
| `docs\` | User and developer guide files. |
| `src\` | Module implementation. |
| `templates\` | HTML report template. |
| `LICENSE` | MIT license. |
| `THIRD-PARTY-NOTICES.md` | Third-party notices. |

## Documentation
- [User guide](docs/GraphConnectivityTests-UserGuide.md) - also `docs/GraphConnectivityTests-UserGuide.html`, a single file to open locally
- [Developer guide](docs/GraphConnectivityTests-Guide.md) - also `docs/GraphConnectivityTests-Guide.html`, a single file to open locally

Project page, releases and change log: https://github.com/Nico77600/GraphConnectivityTests

License: [MIT](LICENSE).

## Disclaimer

This Script is a Personal project.
It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

The administrator checks (`-AdminChecks`) use the administration API of the Exchange Online PowerShell module, which is not documented as a public API and can change. Use `-AllowWrite` only on a test mailbox.
