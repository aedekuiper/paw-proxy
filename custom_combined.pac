// This PAC file will provide proxy config to Microsoft 365 services
//  using data from the public web service for all endpoints
function FindProxyForURL(url, host)
{
    var direct = "DIRECT";
    var proxyServer = "PROXY 127.0.0.2:8080";

        host = host.toLowerCase();
    // Local/intranet bypass
    if (host == "localhost" || shExpMatch(host, "127.*") || isPlainHostName(host)) {
        return "DIRECT";
    }

    // Custom override domains
    if (
        shExpMatch(host, "*.githubusercontent.com") ||
        shExpMatch(host, "*.azure.com") ||
        shExpMatch(host, "*.azure.net") ||
        shExpMatch(host, "*.microsoft.com") ||
        shExpMatch(host, "*.windowsupdate.com") ||
        shExpMatch(host, "*.microsoftonline.com") ||
        shExpMatch(host, "*.microsoftonline.cn") ||
        shExpMatch(host, "*.windows.net") ||
        shExpMatch(host, "*.windowsazure.com") ||
        shExpMatch(host, "*.windowsazure.cn") ||
        shExpMatch(host, "*.azure.cn") ||
        shExpMatch(host, "*.loganalytics.io") ||
        shExpMatch(host, "*.applicationinsights.io") ||
        shExpMatch(host, "*.vsassets.io") ||
        shExpMatch(host, "*.azure-automation.net") ||
        shExpMatch(host, "*.visualstudio.com") ||
        shExpMatch(host, "portal.office.com") ||
        shExpMatch(host, "*.aspnetcdn.com") ||
        shExpMatch(host, "*.sharepointonline.com") ||
        shExpMatch(host, "*.msecnd.net") ||
        shExpMatch(host, "*.msocdn.com") ||
        shExpMatch(host, "*.webtrends.com")
    ) {
        return "DIRECT";
    }


    if(shExpMatch(host, "cdn.odc.officeapps.live.com")
        || shExpMatch(host, "cdn.uci.officeapps.live.com"))
    {
        return proxyServer;
    }

    if(shExpMatch(host, "*.auth.microsoft.com")
        || shExpMatch(host, "*.lync.com")
        || shExpMatch(host, "*.mail.protection.outlook.com")
        || shExpMatch(host, "*.msftidentity.com")
        || shExpMatch(host, "*.msidentity.com")
        || shExpMatch(host, "*.mx.microsoft")
        || shExpMatch(host, "*.officeapps.live.com")
        || shExpMatch(host, "*.online.office.com")
        || shExpMatch(host, "*.protection.office.com")
        || shExpMatch(host, "*.protection.outlook.com")
        || shExpMatch(host, "*.security.microsoft.com")
        || shExpMatch(host, "*.sharepoint.com")
        || shExpMatch(host, "*.teams.cloud.microsoft")
        || shExpMatch(host, "*.teams.microsoft.com")
        || shExpMatch(host, "account.activedirectory.windowsazure.com")
        || shExpMatch(host, "accounts.accesscontrol.windows.net")
        || shExpMatch(host, "adminwebservice.microsoftonline.com")
        || shExpMatch(host, "api.passwordreset.microsoftonline.com")
        || shExpMatch(host, "autologon.microsoftazuread-sso.com")
        || shExpMatch(host, "becws.microsoftonline.com")
        || shExpMatch(host, "ccs.login.microsoftonline.com")
        || shExpMatch(host, "clientconfig.microsoftonline-p.net")
        || shExpMatch(host, "companymanager.microsoftonline.com")
        || shExpMatch(host, "compliance.microsoft.com")
        || shExpMatch(host, "defender.microsoft.com")
        || shExpMatch(host, "device.login.microsoftonline.com")
        || shExpMatch(host, "graph.microsoft.com")
        || shExpMatch(host, "graph.windows.net")
        || shExpMatch(host, "login-us.microsoftonline.com")
        || shExpMatch(host, "login.microsoft.com")
        || shExpMatch(host, "login.microsoftonline-p.com")
        || shExpMatch(host, "login.microsoftonline.com")
        || shExpMatch(host, "login.windows.net")
        || shExpMatch(host, "logincert.microsoftonline.com")
        || shExpMatch(host, "loginex.microsoftonline.com")
        || shExpMatch(host, "nexus.microsoftonline-p.com")
        || shExpMatch(host, "office.live.com")
        || shExpMatch(host, "outlook.cloud.microsoft")
        || shExpMatch(host, "outlook.office.com")
        || shExpMatch(host, "outlook.office365.com")
        || shExpMatch(host, "passwordreset.microsoftonline.com")
        || shExpMatch(host, "protection.office.com")
        || shExpMatch(host, "provisioningapi.microsoftonline.com")
        || shExpMatch(host, "purview.microsoft.com")
        || shExpMatch(host, "security.microsoft.com")
        || shExpMatch(host, "smtp.office365.com")
        || shExpMatch(host, "teams.cloud.microsoft")
        || shExpMatch(host, "teams.microsoft.com"))
    {
        return direct;
    }

    return proxyServer;

}

