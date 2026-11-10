#
#  Graph Connectivity Tests - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 1.1.0
#
#  This file is read by Invoke-GraphConnectivityTests.ps1 and by the window (-Gui). It is a PowerShell
#  data file: text between quotes, $true / $false, numbers, and @( ) for lists. Lines starting with # are
#  comments. Relative paths are relative to the tool folder. Every value is checked at start; all the
#  problems are listed at once. The parameters of the command line override these values.
#
#  No secret here: the sign-in is interactive (sign-in window or device code); an application signs in
#  with a certificate of the certificate store. Replace every contoso value before the first run.
#
@{
    # ---------------------------------------------------------------------
    # What is tested.
    #   SourceMailbox:   the mailbox that signs in and asks for the free/busy. Its domain is also the
    #                    default DNS domain.
    #   TargetMailboxes: the mailboxes whose free/busy is asked (attendees), 20 at most.
    #   OnPremisesHost:  FQDN of the on-premises Availability service (https://<host>/EWS/Exchange.asmx),
    #                    when Autodiscover does not give the right one. Empty = the host Autodiscover gives.
    #   SkipOnPremisesChecks: $true skips the checks of the on-premises endpoint (certificate, OAuth).
    #   Domain:          DomainDns: the domain to check (empty = the domain of the source).
    #   Cloud:           Global | USGov (GCC High) | USGovDoD | China (21Vianet).
    #   HeaderFile:      MessageHeaders: an .eml or .txt file with the headers of a message (only the header
    #                    block is read). -HeaderFile on the command line.
    #   HeaderText:      MessageHeaders: the headers as text, instead of a file. -HeaderText on the command line.
    # ---------------------------------------------------------------------
    Target = @{
        SourceMailbox        = 'alice@contoso.com'
        TargetMailboxes      = @('bob@contoso.com')
        OnPremisesHost       = ''
        SkipOnPremisesChecks = $false
        Domain               = ''
        Cloud                = 'Global'
        HeaderFile           = ''
        HeaderText           = ''
    }

    # ---------------------------------------------------------------------
    # Who signs in.
    #   Context: 'User' (the source user signs in)
    #            or 'Application' (your application with a certificate, for automation: /users/{source}).
    #   ClientId (User): the public client used to sign in.
    #     'MicrosoftOffice'        default: a client of Microsoft present in every tenant and pre-authorized
    #                              for Microsoft Graph - no application to register, no consent.
    #     'GraphCommandLineTools'  the client of Connect-MgGraph: the permissions below need a consent.
    #     <application ID>         your own application (multi-tenant if the users come from several
    #                              tenants), with a redirect URI of type Mobile and desktop.
    #   ClientId (Application): the application ID of your application.
    #   RedirectUri: redirect URI of your application (empty = the default of the client).
    #   Scopes: permissions of Microsoft Graph asked by name (empty = the default of the client).
    #   TenantId: tenant ID or domain (empty = the tenant of the domain of the source). Needed for Application.
    #   CertificateThumbprint (Application): certificate in Cert:\CurrentUser\My or LocalMachine\My.
    #   OnPremisesAuthentication: a source mailbox on-premises is reached with 'OAuth' (hybrid modern
    #     authentication, same sign-in), 'Windows' (current Windows account, computer of the domain) or
    #     'Credential' (an account and its password typed when needed, from any computer).
    #   OnPremisesCredentialScheme (Credential): 'Auto' (the best scheme EWS offers: Negotiate, NTLM, then
    #     Basic), 'Negotiate', 'NTLM' or 'Basic' (Basic only over https://).
    #   OnPremisesUserName (Credential): user name proposed (DOMAIN\user or user@domain; empty = the source
    #     mailbox). The password is never stored: it is typed (Get-Credential, the window) or passed
    #     with -OnPremisesCredential.
    #   AdminClientId: client of the optional administrator checks ('ExchangeOnlinePowerShell', the client of
    #     Connect-ExchangeOnline, or an application ID).
    # ---------------------------------------------------------------------
    Identity = @{
        Context                  = 'User'
        ClientId                 = 'MicrosoftOffice'
        RedirectUri              = ''
        Scopes                   = @()
        TenantId                 = ''
        CertificateThumbprint    = ''
        OnPremisesAuthentication = 'OAuth'
        OnPremisesCredentialScheme = 'Auto'
        OnPremisesUserName       = ''
        AdminClientId            = 'ExchangeOnlinePowerShell'
    }

    # ---------------------------------------------------------------------
    # How it is tested.
    #   DefaultType: FreeBusy | MailboxServices | Provisioning | DomainDns | MessageHeaders | All
    #   SignIn: 'Auto' (the sign-in window when this session can show one, else a device code),
    #           'Window' (Edge or Chrome, temporary profile) or 'DeviceCode' (a code typed on any device).
    #   AllowWrite: $true lets MailboxServices schedule the automatic replies, read them back and restore
    #     them (only when they are disabled). -AllowWrite on the command line.
    #   AdminChecks: $true adds the hybrid configuration of Exchange Online read by an Exchange administrator
    #     (one more sign-in). -AdminChecks on the command line.
    #   FreeBusyDays (1-62), FreeBusyIntervalMinutes: the period asked from today, and the slot length.
    #   SkipDnsCrossCheck: $true leaves out the DNS of today in MessageHeaders (no DNS query at all).
    #     -SkipDnsCrossCheck on the command line.
    #   DnsServer: the DNS servers to ask (IP addresses or names), for DomainDns and the DNS of today of
    #     MessageHeaders; empty: the DNS servers of Windows. -DnsServer on the command line.
    #   DnsProbeName: a name that always exists, asked first by MessageHeaders. When no DNS server answers it,
    #     the header is read offline (no comparison with the DNS of today) instead of failing.
    # ---------------------------------------------------------------------
    Test = @{
        DefaultType             = 'FreeBusy'
        SignIn                  = 'Auto'
        AllowWrite              = $false
        AdminChecks             = $false
        FreeBusyDays            = 7
        FreeBusyIntervalMinutes = 30
        OAuthPollTimeoutSeconds = 600
        HttpTimeoutSeconds      = 60
        CertificateWarningDays  = 30
        MaxRetries              = 2
        UserAgent               = 'GraphConnectivityTests/1.0'
        SkipDnsCrossCheck       = $false
        DnsServer               = @()
        DnsProbeName            = 'microsoft.com'
    }

    # ---------------------------------------------------------------------
    # Report files: one folder per run under OutputPath.
    # ---------------------------------------------------------------------
    Report = @{
        OutputPath   = 'reports'
        FilePrefix   = 'GraphConnectivityTests'
        Formats      = @('Csv', 'Html')
        CsvDelimiter = ';'
    }

    Logging = @{
        Path          = 'logs'
        RetentionDays = 14
    }
}
