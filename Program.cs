const string ProductName = "OpenInternetExplorer";
const string Prefix_HTTPS = "https://";
const string Prefix_HTTP = "http://";

try
{
    var url = GetStartPage(args);
    try
    {
        OpenInternetExplorer(url);
    }
    catch
    {
        // https://stackoverflow.com/questions/56044878/how-to-fix-an-hresult-0x8150002e-exception
        TryKillInternetExplorer();
        OpenInternetExplorer(url);
    }
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
        if (Environment.Is64BitOperatingSystem)
        {
            url = GetStartPageByRegistry(RegistryView.Registry64);
            if (IsHttpUrl(url, httpsOnly)) return url;
        }
        url = GetStartPageByRegistry(RegistryView.Registry32);
        if (IsHttpUrl(url, httpsOnly)) return url;
    }
    catch
    {

    }
    return "https://www.bing.com";
}

static string? GetStartPageByRegistry(RegistryView registryView)
{
    using var registryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, registryView)
        .OpenSubKey(@"Software\Microsoft\Internet Explorer\Main");
    return registryKey.GetValue("Start Page")?.ToString();
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

static void OpenInternetExplorer(string url)
{
    // https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752084(v=vs.85)
    var IE = new InternetExplorer
    {
        Visible = true
    };
    IE.Navigate(url);
}

static void TryKillInternetExplorer()
{
    var processes = Process.GetProcessesByName("iexplore");
    foreach (var process in processes)
    {
        try
        {
            process.Kill();
        }
        catch
        {

        }
    }
}