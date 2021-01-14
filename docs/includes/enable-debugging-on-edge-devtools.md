如果加载项正在 Microsoft Edge 中运行，则无 UI 的代码将无法默认附加到调试程序。
无 UI 的代码是指当任务窗格不可见时运行的任何代码，例如加载项命令。 要启用调试，需要运行以下 [Windows PowerShell](/powershell/scripting/getting-started/getting-started-with-windows-powershell) 命令。

1. 请运行以下命令，获取有关 **Microsoft.Win32WebViewHost** 应用包的信息。
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    该命令会列出与以下输出类似的应用包信息。
    
    ```powershell
    Name              : Microsoft.Win32WebViewHost
    Publisher         : CN=Microsoft Windows, O=Microsoft Corporation, L=Redmond, S=Washington, C=US
    Architecture      : Neutral
    ResourceId        : neutral
    Version           : 10.0.18362.449
    PackageFullName   : Microsoft.Win32WebViewHost_10.0.18362.449_neutral_neutral_cw5n1h2txyewy
    InstallLocation   : C:\Windows\SystemApps\Microsoft.Win32WebViewHost_cw5n1h2txyewy
    IsFramework       : False
    PackageFamilyName : Microsoft.Win32WebViewHost_cw5n1h2txyewy
    PublisherId       : cw5n1h2txyewy
    IsResourcePackage : False
    IsBundle          : False
    IsDevelopmentMode : False
    NonRemovable      : True
    IsPartiallyStaged : False
    SignatureKind     : System
    Status            : Ok
    ```
    
2. 请运行以下命令来启用调试。 使用从上一命令列出的 **PackageFullName** 的值。
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. 如果 Office 已在运行，请关闭再重启 Office，使其获取调试更改。