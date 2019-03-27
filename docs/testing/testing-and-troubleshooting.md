---
title: 排查 Office 加载项中的用户错误
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 84f18543c7bafac905805095c89f8e19a855ea76
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871071"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="0102f-102">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="0102f-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="0102f-p101">有时，您的用户在使用您开发的 Office 外接程序时可能会遇到问题。例如，外接程序无法加载或无法访问。使用本文中的信息有助于解决您的用户在使用 Office 外接程序时遇到的常见问题。</span><span class="sxs-lookup"><span data-stu-id="0102f-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="0102f-106">还可以使用 [Fiddler](https://www.telerik.com/fiddler) 发现和调试加载项问题。</span><span class="sxs-lookup"><span data-stu-id="0102f-106">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

<span data-ttu-id="0102f-107">解决用户的问题后，可以[在 AppSource 中直接回复客户评论](/office/dev/store/create-effective-office-store-listings)。</span><span class="sxs-lookup"><span data-stu-id="0102f-107">After you resolve the user's issue, you can [respond directly to customer reviews in AppSource](/office/dev/store/create-effective-office-store-listings).</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="0102f-108">常见错误和故障排除步骤</span><span class="sxs-lookup"><span data-stu-id="0102f-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="0102f-109">下表列出了用户可能遇到的常见错误消息以及用户可以采取以解决这些错误的步骤。</span><span class="sxs-lookup"><span data-stu-id="0102f-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="0102f-110">**错误消息**</span><span class="sxs-lookup"><span data-stu-id="0102f-110">**Error message**</span></span>|<span data-ttu-id="0102f-111">**解决方案**</span><span class="sxs-lookup"><span data-stu-id="0102f-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="0102f-112">应用错误：无法访问目录</span><span class="sxs-lookup"><span data-stu-id="0102f-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="0102f-p102">验证防火墙设置。“目录”是指 AppSource。此消息表明用户无法访问 AppSource。</span><span class="sxs-lookup"><span data-stu-id="0102f-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="0102f-p103">应用错误：无法启动此应用。若要忽略此问题，请关闭这个对话框。若要重试，请单击“重启”。</span><span class="sxs-lookup"><span data-stu-id="0102f-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="0102f-117">确认已安装最新的 Office 更新，或下载 [Office 2013 更新](https://support.microsoft.com/kb/2986156/)。</span><span class="sxs-lookup"><span data-stu-id="0102f-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="0102f-118">错误：对象不 支持此属性或方法 "defineProperty"</span><span class="sxs-lookup"><span data-stu-id="0102f-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="0102f-p104">确认 Internet Explorer 不是在兼容模式下运行。转到“工具”>“**兼容性视图设置**”。</span><span class="sxs-lookup"><span data-stu-id="0102f-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="0102f-p105">很抱歉，我们无法加载 该应用程序，因为您的浏览器 版本不受支持。 单击此处查看 支持的浏览器版本的列表。</span><span class="sxs-lookup"><span data-stu-id="0102f-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="0102f-p106">确保浏览器支持 HTML5 本地存储，或重置您的 Internet Explorer 设置。有关受支持的浏览器的信息，请参阅 [运行 Office 加载项的要求](../concepts/requirements-for-running-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="0102f-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|


## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="0102f-125">Outlook 外接程序不能正常工作</span><span class="sxs-lookup"><span data-stu-id="0102f-125">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="0102f-126">如果在 Windows 上运行的 Outlook 外接程序不能正常工作，请尝试在 Internet Explorer 中启用脚本调试。</span><span class="sxs-lookup"><span data-stu-id="0102f-126">If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="0102f-127">转到“工具”>“**Internet 选项**” > “**高级**”。</span><span class="sxs-lookup"><span data-stu-id="0102f-127">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="0102f-128">在“**浏览**”下，取消选中“**禁用脚本调试 (Internet Explorer)**”和“**禁用脚本调试 (其他)**”。</span><span class="sxs-lookup"><span data-stu-id="0102f-128">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="0102f-p107">我们建议仅在解决问题时取消选中这些设置。如果你将其保持未选中状态，你在浏览时会收到提示。解决此问题后，再次选中“**禁用脚本调试(Internet Explorer)**”和“**禁用脚本调试(其他)**”。</span><span class="sxs-lookup"><span data-stu-id="0102f-p107">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="0102f-132">外接程序在 Office 2013 中无法激活</span><span class="sxs-lookup"><span data-stu-id="0102f-132">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="0102f-133">如果在用户执行下列步骤时外接程序无法激活：</span><span class="sxs-lookup"><span data-stu-id="0102f-133">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="0102f-134">使用 Microsoft 帐户在 Office 2013 中登录。</span><span class="sxs-lookup"><span data-stu-id="0102f-134">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="0102f-135">为其 Microsoft 帐户启用两步验证。</span><span class="sxs-lookup"><span data-stu-id="0102f-135">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="0102f-136">尝试插入外接程序时在收到提示的时候验证其身份。</span><span class="sxs-lookup"><span data-stu-id="0102f-136">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="0102f-137">确认是否已安装最新的 Office 更新程序，或下载 [Office 2013 更新程序](https://support.microsoft.com/kb/2986156/)。</span><span class="sxs-lookup"><span data-stu-id="0102f-137">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="0102f-138">外接程序无法在任务窗格中加载，或外接程序清单存在其他问题</span><span class="sxs-lookup"><span data-stu-id="0102f-138">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="0102f-139">请参阅[验证并排查清单问题](troubleshoot-manifest.md)，针对外接程序清单问题进行调试。</span><span class="sxs-lookup"><span data-stu-id="0102f-139">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="0102f-140">无法显示外接程序对话框</span><span class="sxs-lookup"><span data-stu-id="0102f-140">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="0102f-p108">使用 Office 外接程序时，将要求用户允许显示对话框。用户选择“**允许**”，将出现以下错误消息：</span><span class="sxs-lookup"><span data-stu-id="0102f-p108">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="0102f-p109">“浏览器中的安全设置阻止创建对话框。请尝试使用其他浏览器，或者配置浏览器，使地址栏中显示的 [URL] 和域处于同一安全区域。”</span><span class="sxs-lookup"><span data-stu-id="0102f-p109">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![对话框错误消息的屏幕截图](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="0102f-146">**受影响的浏览器**</span><span class="sxs-lookup"><span data-stu-id="0102f-146">**Affected browsers**</span></span>|<span data-ttu-id="0102f-147">**受影响的平台**</span><span class="sxs-lookup"><span data-stu-id="0102f-147">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="0102f-148">Internet Explorer、Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="0102f-148">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="0102f-149">Office Online</span><span class="sxs-lookup"><span data-stu-id="0102f-149">Office Online</span></span>|

<span data-ttu-id="0102f-p110">若要解决此问题，最终用户或管理员可以向 Internet Explorer 中的受信任站点列表添加外接程序的域。无论使用的是 Internet Explorer 还是 Microsoft Edge 浏览器，请使用相同过程。</span><span class="sxs-lookup"><span data-stu-id="0102f-p110">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0102f-152">请勿将不信任的加载项的 URL 添加到受信任网站列表中。</span><span class="sxs-lookup"><span data-stu-id="0102f-152">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="0102f-153">要将 URL 添加到受信任站点的列表中，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="0102f-153">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="0102f-154">在 Internet Explorer 中，选择“工具”按钮，然后转到“**Internet 选项**” > “**安全**”。</span><span class="sxs-lookup"><span data-stu-id="0102f-154">In Internet Explorer, choose the Tools button, and go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="0102f-155">选择“**受信任站点**”区域，并选择“**网站**”。</span><span class="sxs-lookup"><span data-stu-id="0102f-155">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="0102f-156">输入错误消息中显示的 URL，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="0102f-156">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="0102f-p111">再次尝试使用外接程序。如果问题仍然存在，请验证其他安全区域的设置，并确保外接程序域与 Office 应用程序地址栏中显示的 URL 处于同一区域。</span><span class="sxs-lookup"><span data-stu-id="0102f-p111">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="0102f-p112">在弹出模式中使用对话框 API 时，会出现此问题。若要避免出现此问题，请使用 [displayInFrame](/javascript/api/office/office.ui) 标记。这要求页面支持在 iframe 中进行显示。以下示例演示如何使用此标记。</span><span class="sxs-lookup"><span data-stu-id="0102f-p112">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js

Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="0102f-163">对加载项命令（包括功能区按钮和菜单项）的更改未生效</span><span class="sxs-lookup"><span data-stu-id="0102f-163">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>
<span data-ttu-id="0102f-p113">有时，对加载项命令（如功能区按钮图标或菜单项文本）的更改似乎未生效。请清除旧版的 Office 缓存。</span><span class="sxs-lookup"><span data-stu-id="0102f-p113">Sometimes changes to add-in commands such as the icon for a ribbon button or the text of a menu item do not seem to take effect. Clear the Office cache of the old versions.</span></span>

#### <a name="for-windows"></a><span data-ttu-id="0102f-166">对于 Windows：</span><span class="sxs-lookup"><span data-stu-id="0102f-166">For Windows:</span></span>
<span data-ttu-id="0102f-167">删除文件夹 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 的内容。</span><span class="sxs-lookup"><span data-stu-id="0102f-167">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="0102f-168">对于 Mac：</span><span class="sxs-lookup"><span data-stu-id="0102f-168">For Mac:</span></span>
<span data-ttu-id="0102f-169">删除文件夹 `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` 的内容。</span><span class="sxs-lookup"><span data-stu-id="0102f-169">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="0102f-170">对于 iOS：</span><span class="sxs-lookup"><span data-stu-id="0102f-170">For iOS:</span></span>
<span data-ttu-id="0102f-p114">在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="0102f-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="0102f-173">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0102f-173">See also</span></span>

- [<span data-ttu-id="0102f-174">在 Office Online 中调试加载项</span><span class="sxs-lookup"><span data-stu-id="0102f-174">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="0102f-175">将 Office 外接程序旁加载到 iPad 和 Mac 上</span><span class="sxs-lookup"><span data-stu-id="0102f-175">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="0102f-176">在 iPad 和 Mac 上调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="0102f-176">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="0102f-177">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="0102f-177">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
    
