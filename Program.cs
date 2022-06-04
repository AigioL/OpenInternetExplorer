const string ProductName = "OpenInternetExplorer";
const string Prefix_HTTPS = "https://";
const string Prefix_HTTP = "http://";

try
{
    var obj = new ScriptControl
    {
        Language = "vbscript",
    };
    // https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.CreateObject#example
    const string content = $"Set Web = CreateObject(\"InternetExplorer.Application\")\r\nWeb.Visible = True\r\nWeb.Navigate \"{{0}}\"";
    obj.ExecuteStatement(string.Format(content, GetStartPage(args)));
}
catch (Exception ex)
{
    MessageBox.Show(ex.ToString(), ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
}

static bool IsHttpUrl([NotNullWhen(true)] string? url, bool httpsOnly = false) => url != null &&
       (url.StartsWith(Prefix_HTTPS, StringComparison.OrdinalIgnoreCase) ||
             (!httpsOnly && url.StartsWith(Prefix_HTTP, StringComparison.OrdinalIgnoreCase)));

static string GetStartPage(string[] args)
{
    var url = GetArgument(args, 0);
    var httpsOnly = GetArgumentB(args, 1);
    if (IsHttpUrl(url, httpsOnly)) return url;
    try
    {
        using var registryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,
             Environment.Is64BitOperatingSystem ?
                 RegistryView.Registry64 :
                 RegistryView.Registry32)
             .OpenSubKey(@"Software\Microsoft\Internet Explorer\Main");
        url = registryKey.GetValue("Start Page")?.ToString();
        if (IsHttpUrl(url, httpsOnly)) return url;
    }
    catch
    {

    }
    return "https://www.bing.com";
}

static string? GetArgument(string[] args, int index)
{
    try
    {
        return args[index];
    }
    catch
    {
        return null;
    }
}

static bool GetArgumentB(string[] args, int index, bool defaultValue = false)
{
    try
    {
        return bool.Parse(args[index]);
    }
    catch
    {
    }
    return defaultValue;
}