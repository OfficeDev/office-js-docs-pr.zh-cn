---
title: 排查 Office 加载项中的用户错误
description: ''
ms.date: 12/31/2019
localization_priority: Normal
ms.openlocfilehash: 27755c6e378c8002565dc1ea938776c66c204c89
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325029"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="68c1d-102">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="68c1d-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="68c1d-p101">有时，您的用户在使用您开发的 Office 外接程序时可能会遇到问题。例如，外接程序无法加载或无法访问。使用本文中的信息有助于解决您的用户在使用 Office 外接程序时遇到的常见问题。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="68c1d-106">还可以使用 [Fiddler](https://www.telerik.com/fiddler) 发现和调试加载项问题。</span><span class="sxs-lookup"><span data-stu-id="68c1d-106">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="68c1d-107">常见错误和故障排除步骤</span><span class="sxs-lookup"><span data-stu-id="68c1d-107">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="68c1d-108">下表列出了用户可能遇到的常见错误消息以及用户可以采取以解决这些错误的步骤。</span><span class="sxs-lookup"><span data-stu-id="68c1d-108">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="68c1d-109">**错误消息**</span><span class="sxs-lookup"><span data-stu-id="68c1d-109">**Error message**</span></span>|<span data-ttu-id="68c1d-110">**解决方案**</span><span class="sxs-lookup"><span data-stu-id="68c1d-110">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="68c1d-111">应用错误：无法访问目录</span><span class="sxs-lookup"><span data-stu-id="68c1d-111">App error: Catalog could not be reached</span></span>|<span data-ttu-id="68c1d-p102">验证防火墙设置。“目录”是指 AppSource。此消息表明用户无法访问 AppSource。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="68c1d-p103">应用错误：无法启动此应用。若要忽略此问题，请关闭这个对话框。若要重试，请单击“重启”。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="68c1d-116">确认已安装最新的 Office 更新，或下载 [Office 2013 更新](https://support.microsoft.com/kb/2986156/)。</span><span class="sxs-lookup"><span data-stu-id="68c1d-116">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="68c1d-117">错误：对象不 支持此属性或方法 "defineProperty"</span><span class="sxs-lookup"><span data-stu-id="68c1d-117">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="68c1d-p104">确认 Internet Explorer 未在兼容模式下运行。转到 "工具" >**兼容性视图设置**。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools > **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="68c1d-p105">很抱歉，我们无法加载 该应用程序，因为您的浏览器 版本不受支持。 单击此处查看 支持的浏览器版本的列表。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="68c1d-p106">确保浏览器支持 HTML5 本地存储，或重置您的 Internet Explorer 设置。有关受支持的浏览器的信息，请参阅 [运行 Office 加载项的要求](../concepts/requirements-for-running-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="68c1d-124">安装加载项时，状态栏中会显示“加载加载项时出错”</span><span class="sxs-lookup"><span data-stu-id="68c1d-124">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="68c1d-125">关闭 Office。</span><span class="sxs-lookup"><span data-stu-id="68c1d-125">Close Office.</span></span>
2. <span data-ttu-id="68c1d-126">验证清单是否有效</span><span class="sxs-lookup"><span data-stu-id="68c1d-126">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="68c1d-127">重启加载项</span><span class="sxs-lookup"><span data-stu-id="68c1d-127">Restart the add-in</span></span>
4. <span data-ttu-id="68c1d-128">再次安装加载项。</span><span class="sxs-lookup"><span data-stu-id="68c1d-128">Install the add-in again.</span></span>

<span data-ttu-id="68c1d-129">你还可以向我们提供反馈：如果使用 Windows 版 Excel 或 Mac 版 Excel，可以直接从 Excel 向 Office 扩展性团队报告反馈。</span><span class="sxs-lookup"><span data-stu-id="68c1d-129">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="68c1d-130">若要执行此操作，请选择“**文件**” | “**反馈**” | “**发送哭脸**”。</span><span class="sxs-lookup"><span data-stu-id="68c1d-130">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="68c1d-131">发送哭脸将提供必要的日志，以帮助我们了解该问题。</span><span class="sxs-lookup"><span data-stu-id="68c1d-131">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="68c1d-132">Outlook 外接程序不能正常工作</span><span class="sxs-lookup"><span data-stu-id="68c1d-132">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="68c1d-133">如果在 Windows 上运行并[使用 Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) 的 Outlook 加载项不能正常工作，请尝试在 Internet Explorer 中启用脚本调试。</span><span class="sxs-lookup"><span data-stu-id="68c1d-133">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="68c1d-134">转到 Tools > **Internet Options** > **Advanced**"。</span><span class="sxs-lookup"><span data-stu-id="68c1d-134">Go to Tools > **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="68c1d-135">在“浏览”\*\*\*\* 下，取消选中“禁用脚本调试 (Internet Explorer)”\*\*\*\* 和“禁用脚本调试 (其他)”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="68c1d-135">Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="68c1d-p108">我们建议您仅在解决问题时取消选中这些设置。如果您将其保持未选中状态，您在浏览时将会收到提示。解决此问题后，再次选中“禁用脚本调试 (Internet Explorer)”\*\*\*\* 和“禁用脚本调试 (其他)”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p108">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="68c1d-139">外接程序在 Office 2013 中无法激活</span><span class="sxs-lookup"><span data-stu-id="68c1d-139">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="68c1d-140">如果在用户执行下列步骤时外接程序无法激活：</span><span class="sxs-lookup"><span data-stu-id="68c1d-140">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="68c1d-141">使用 Microsoft 帐户在 Office 2013 中登录。</span><span class="sxs-lookup"><span data-stu-id="68c1d-141">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="68c1d-142">为其 Microsoft 帐户启用两步验证。</span><span class="sxs-lookup"><span data-stu-id="68c1d-142">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="68c1d-143">尝试插入外接程序时在收到提示的时候验证其身份。</span><span class="sxs-lookup"><span data-stu-id="68c1d-143">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="68c1d-144">确认是否已安装最新的 Office 更新程序，或下载 [Office 2013 更新程序](https://support.microsoft.com/kb/2986156/)。</span><span class="sxs-lookup"><span data-stu-id="68c1d-144">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="68c1d-145">外接程序无法在任务窗格中加载，或外接程序清单存在其他问题</span><span class="sxs-lookup"><span data-stu-id="68c1d-145">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="68c1d-146">请参阅[验证 Office 加载项的清单](troubleshoot-manifest.md)和[使用运行时日志记录功能调试加载项](runtime-logging.md)，以调试加载项清单问题。</span><span class="sxs-lookup"><span data-stu-id="68c1d-146">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="68c1d-147">无法显示外接程序对话框</span><span class="sxs-lookup"><span data-stu-id="68c1d-147">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="68c1d-p109">使用 Office 外接程序时，将要求用户允许显示对话框。用户选择“**允许**”，将出现以下错误消息：</span><span class="sxs-lookup"><span data-stu-id="68c1d-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="68c1d-p110">“浏览器中的安全设置阻止创建对话框。请尝试使用其他浏览器，或者配置浏览器，使地址栏中显示的 [URL] 和域处于同一安全区域。”</span><span class="sxs-lookup"><span data-stu-id="68c1d-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![对话框错误消息的屏幕截图](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="68c1d-153">**受影响的浏览器**</span><span class="sxs-lookup"><span data-stu-id="68c1d-153">**Affected browsers**</span></span>|<span data-ttu-id="68c1d-154">**受影响的平台**</span><span class="sxs-lookup"><span data-stu-id="68c1d-154">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="68c1d-155">Internet Explorer、Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="68c1d-155">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="68c1d-156">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="68c1d-156">Office on the web</span></span>|

<span data-ttu-id="68c1d-p111">若要解决此问题，最终用户或管理员可以向 Internet Explorer 中的受信任站点列表添加外接程序的域。无论使用的是 Internet Explorer 还是 Microsoft Edge 浏览器，请使用相同过程。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="68c1d-159">请勿将不信任的加载项的 URL 添加到受信任网站列表中。</span><span class="sxs-lookup"><span data-stu-id="68c1d-159">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="68c1d-160">要将 URL 添加到受信任站点的列表中，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="68c1d-160">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="68c1d-161">在“**控制面板**”中，转到“**Internet 选项**” > “**安全性**”。</span><span class="sxs-lookup"><span data-stu-id="68c1d-161">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="68c1d-162">选择“**受信任站点**”区域，并选择“**网站**”。</span><span class="sxs-lookup"><span data-stu-id="68c1d-162">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="68c1d-163">输入错误消息中显示的 URL，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="68c1d-163">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="68c1d-p112">再次尝试使用外接程序。如果问题仍然存在，请验证其他安全区域的设置，并确保外接程序域与 Office 应用程序地址栏中显示的 URL 处于同一区域。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="68c1d-p113">在弹出模式中使用对话框 API 时，会出现此问题。若要避免出现此问题，请使用 [displayInFrame](/javascript/api/office/office.ui) 标记。这要求页面支持在 iframe 中进行显示。以下示例演示如何使用此标记。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="68c1d-170">对加载项命令（包括功能区按钮和菜单项）的更改未生效</span><span class="sxs-lookup"><span data-stu-id="68c1d-170">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="68c1d-171">如果在清单中进行的更改（如功能区按钮图标的文件名或菜单项的文本）似乎没有生效，请尝试清除计算机上的 Office 缓存。</span><span class="sxs-lookup"><span data-stu-id="68c1d-171">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="68c1d-172">对于 Windows：</span><span class="sxs-lookup"><span data-stu-id="68c1d-172">For Windows:</span></span>
<span data-ttu-id="68c1d-173">删除文件夹 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 的内容。</span><span class="sxs-lookup"><span data-stu-id="68c1d-173">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="68c1d-174">对于 Mac：</span><span class="sxs-lookup"><span data-stu-id="68c1d-174">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="68c1d-175">对于 iOS：</span><span class="sxs-lookup"><span data-stu-id="68c1d-175">For iOS:</span></span>
<span data-ttu-id="68c1d-p114">在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="68c1d-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="68c1d-178">对静态文件（例如 JavaScript、HTML 和 CSS）的更改未生效</span><span class="sxs-lookup"><span data-stu-id="68c1d-178">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="68c1d-179">浏览器可能正在缓存这些文件。</span><span class="sxs-lookup"><span data-stu-id="68c1d-179">The browser may be caching these files.</span></span> <span data-ttu-id="68c1d-180">若要阻止此操作，请在开发时关闭客户端缓存。</span><span class="sxs-lookup"><span data-stu-id="68c1d-180">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="68c1d-181">详细信息取决于你所使用的服务器类型。</span><span class="sxs-lookup"><span data-stu-id="68c1d-181">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="68c1d-182">在大多数情况下，它涉及将某些标头添加到 HTTP 响应。</span><span class="sxs-lookup"><span data-stu-id="68c1d-182">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="68c1d-183">我们建议使用以下设置：</span><span class="sxs-lookup"><span data-stu-id="68c1d-183">We suggest the following set:</span></span>

- <span data-ttu-id="68c1d-184">Cache-Control：“private、no-cache、no-store”</span><span class="sxs-lookup"><span data-stu-id="68c1d-184">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="68c1d-185">Pragma：“no-cache”</span><span class="sxs-lookup"><span data-stu-id="68c1d-185">Pragma: "no-cache"</span></span>
- <span data-ttu-id="68c1d-186">过期：“-1”</span><span class="sxs-lookup"><span data-stu-id="68c1d-186">Expires: "-1"</span></span>

<span data-ttu-id="68c1d-187">有关在 Node.JS Express 服务器中执行此操作的示例，请参阅[此 app.js 文件](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)。</span><span class="sxs-lookup"><span data-stu-id="68c1d-187">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="68c1d-188">有关 ASP.NET 项目中的示例，请参阅[此 cshtml 文件](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/src/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)。</span><span class="sxs-lookup"><span data-stu-id="68c1d-188">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/src/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="68c1d-189">如果加载项托管在 Internet Information Server (IIS) 中，则还可以将以下内容添加到 web.config 中。</span><span class="sxs-lookup"><span data-stu-id="68c1d-189">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="68c1d-190">如果这些步骤一开始似乎不起作用，则可能需要清除浏览器的缓存。</span><span class="sxs-lookup"><span data-stu-id="68c1d-190">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="68c1d-191">请通过浏览器的 UI 执行此操作。</span><span class="sxs-lookup"><span data-stu-id="68c1d-191">Do this through the UI of the browser.</span></span> <span data-ttu-id="68c1d-192">有时，当你尝试在边缘 UI 中清除边缘缓存时，无法成功清除它。</span><span class="sxs-lookup"><span data-stu-id="68c1d-192">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="68c1d-193">如果出现这种情况，请在 Windows 命令提示符中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="68c1d-193">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a><span data-ttu-id="68c1d-194">另请参阅</span><span class="sxs-lookup"><span data-stu-id="68c1d-194">See also</span></span>

- [<span data-ttu-id="68c1d-195">在 Office 网页版中调试加载项</span><span class="sxs-lookup"><span data-stu-id="68c1d-195">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="68c1d-196">将 Office 外接程序旁加载到 iPad 和 Mac 上</span><span class="sxs-lookup"><span data-stu-id="68c1d-196">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="68c1d-197">在 iPad 和 Mac 上调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="68c1d-197">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="68c1d-198">验证 Office 加载项的清单</span><span class="sxs-lookup"><span data-stu-id="68c1d-198">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="68c1d-199">使用运行时日志记录功能调试加载项</span><span class="sxs-lookup"><span data-stu-id="68c1d-199">Debug your add-in with runtime logging</span></span>](runtime-logging.md)