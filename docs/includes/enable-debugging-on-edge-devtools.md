<span data-ttu-id="fb7de-101">如果加载项正在 Microsoft Edge 中运行，则无 UI 的代码将无法默认附加到调试程序。</span><span class="sxs-lookup"><span data-stu-id="fb7de-101">When the add-in is running in Microsoft Edge, UI-less code will not be able to attach to a debugger by default.</span></span>
<span data-ttu-id="fb7de-102">无 UI 的代码是指当任务窗格不可见时运行的任何代码，例如加载项命令。</span><span class="sxs-lookup"><span data-stu-id="fb7de-102">UI-less code is any code running while the task pane is not visible, such as add-in commands.</span></span> <span data-ttu-id="fb7de-103">要启用调试，需要运行以下 [Windows PowerShell](https://docs.microsoft.com/powershell/scripting/getting-started/getting-started-with-windows-powershell) 命令。</span><span class="sxs-lookup"><span data-stu-id="fb7de-103">To enable debugging, you need to run the following [Windows PowerShell](https://docs.microsoft.com/powershell/scripting/getting-started/getting-started-with-windows-powershell) commands.</span></span>

1. <span data-ttu-id="fb7de-104">请运行以下命令，获取有关 **Microsoft.Win32WebViewHost** 应用包的信息。</span><span class="sxs-lookup"><span data-stu-id="fb7de-104">Run the following command to get information for the **Microsoft.Win32WebViewHost** app package.</span></span>
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    <span data-ttu-id="fb7de-105">该命令会列出与以下输出类似的应用包信息。</span><span class="sxs-lookup"><span data-stu-id="fb7de-105">The command lists app package information similar to the following output.</span></span>
    
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
    
2. <span data-ttu-id="fb7de-106">请运行以下命令来启用调试。</span><span class="sxs-lookup"><span data-stu-id="fb7de-106">Run the following command to verify content filtering is enabled:</span></span> <span data-ttu-id="fb7de-107">使用从上一命令列出的 **PackageFullName** 的值。</span><span class="sxs-lookup"><span data-stu-id="fb7de-107">Use the value for the **PackageFullName** listed from the previous command.</span></span>
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. <span data-ttu-id="fb7de-108">如果 Office 已在运行，请关闭再重启 Office，使其获取调试更改。</span><span class="sxs-lookup"><span data-stu-id="fb7de-108">If Office was already running, close and restart Office so that it picks up the debugging change.</span></span>