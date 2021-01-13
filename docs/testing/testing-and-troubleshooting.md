---
title: 排查 Office 加载项中的用户错误
description: 了解如何排查 Office 加载项中的用户错误。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: e1cb5e0bb8005f04425a5ad9c7e807d10f054e35
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49840089"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="a7bae-103">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="a7bae-103">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="a7bae-p101">有时，您的用户在使用您开发的 Office 外接程序时可能会遇到问题。例如，外接程序无法加载或无法访问。使用本文中的信息有助于解决您的用户在使用 Office 外接程序时遇到的常见问题。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="a7bae-107">还可以使用 [Fiddler](https://www.telerik.com/fiddler) 发现和调试加载项问题。</span><span class="sxs-lookup"><span data-stu-id="a7bae-107">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="a7bae-108">常见错误和故障排除步骤</span><span class="sxs-lookup"><span data-stu-id="a7bae-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="a7bae-109">下表列出了用户可能遇到的常见错误消息以及用户可以采取以解决这些错误的步骤。</span><span class="sxs-lookup"><span data-stu-id="a7bae-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="a7bae-110">**错误消息**</span><span class="sxs-lookup"><span data-stu-id="a7bae-110">**Error message**</span></span>|<span data-ttu-id="a7bae-111">**解决方案**</span><span class="sxs-lookup"><span data-stu-id="a7bae-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="a7bae-112">应用错误：无法访问目录</span><span class="sxs-lookup"><span data-stu-id="a7bae-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="a7bae-p102">验证防火墙设置。“目录”是指 AppSource。此消息表明用户无法访问 AppSource。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="a7bae-p103">应用错误：无法启动此应用。若要忽略此问题，请关闭这个对话框。若要重试，请单击“重启”。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="a7bae-117">确认已安装最新的 Office 更新，或下载 [Office 2013 更新](https://support.microsoft.com/kb/2986156/)。</span><span class="sxs-lookup"><span data-stu-id="a7bae-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="a7bae-118">错误：对象不 支持此属性或方法 "defineProperty"</span><span class="sxs-lookup"><span data-stu-id="a7bae-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="a7bae-119">确认 Internet Explorer 不是在兼容模式下运行。</span><span class="sxs-lookup"><span data-stu-id="a7bae-119">Confirm that Internet Explorer is not running in Compatibility Mode.</span></span> <span data-ttu-id="a7bae-120">转到“工具”>“兼容性视图设置”。</span><span class="sxs-lookup"><span data-stu-id="a7bae-120">Go to Tools > **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="a7bae-p105">很抱歉，我们无法加载 该应用程序，因为您的浏览器 版本不受支持。 单击此处查看 支持的浏览器版本的列表。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="a7bae-p106">确保浏览器支持 HTML5 本地存储，或重置您的 Internet Explorer 设置。有关受支持的浏览器的信息，请参阅 [运行 Office 加载项的要求](../concepts/requirements-for-running-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="a7bae-125">安装加载项时，状态栏中会显示“加载加载项时出错”</span><span class="sxs-lookup"><span data-stu-id="a7bae-125">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="a7bae-126">关闭 Office。</span><span class="sxs-lookup"><span data-stu-id="a7bae-126">Close Office.</span></span>
2. <span data-ttu-id="a7bae-127">验证清单是否有效</span><span class="sxs-lookup"><span data-stu-id="a7bae-127">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="a7bae-128">重启加载项</span><span class="sxs-lookup"><span data-stu-id="a7bae-128">Restart the add-in</span></span>
4. <span data-ttu-id="a7bae-129">再次安装加载项。</span><span class="sxs-lookup"><span data-stu-id="a7bae-129">Install the add-in again.</span></span>

<span data-ttu-id="a7bae-130">你还可以向我们提供反馈：如果使用 Windows 版 Excel 或 Mac 版 Excel，可以直接从 Excel 向 Office 扩展性团队报告反馈。</span><span class="sxs-lookup"><span data-stu-id="a7bae-130">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="a7bae-131">若要执行此操作，请选择“**文件**” | “**反馈**” | “**发送哭脸**”。</span><span class="sxs-lookup"><span data-stu-id="a7bae-131">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="a7bae-132">发送哭脸将提供必要的日志，以帮助我们了解该问题。</span><span class="sxs-lookup"><span data-stu-id="a7bae-132">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="a7bae-133">Outlook 外接程序不能正常工作</span><span class="sxs-lookup"><span data-stu-id="a7bae-133">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="a7bae-134">如果在 Windows 上运行并[使用 Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) 的 Outlook 加载项不能正常工作，请尝试在 Internet Explorer 中启用脚本调试。</span><span class="sxs-lookup"><span data-stu-id="a7bae-134">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="a7bae-135">转到"工具> Internet **选项**  >  **高级"。**</span><span class="sxs-lookup"><span data-stu-id="a7bae-135">Go to Tools > **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="a7bae-136">在“浏览”下，取消选中“禁用脚本调试 (Internet Explorer)”和“禁用脚本调试 (其他)”。</span><span class="sxs-lookup"><span data-stu-id="a7bae-136">Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="a7bae-p108">我们建议您仅在解决问题时取消选中这些设置。如果您将其保持未选中状态，您在浏览时将会收到提示。解决此问题后，再次选中“禁用脚本调试 (Internet Explorer)”和“禁用脚本调试 (其他)”。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p108">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="a7bae-140">外接程序在 Office 2013 中无法激活</span><span class="sxs-lookup"><span data-stu-id="a7bae-140">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="a7bae-141">如果在用户执行下列步骤时外接程序无法激活：</span><span class="sxs-lookup"><span data-stu-id="a7bae-141">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="a7bae-142">使用 Microsoft 帐户在 Office 2013 中登录。</span><span class="sxs-lookup"><span data-stu-id="a7bae-142">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="a7bae-143">为其 Microsoft 帐户启用两步验证。</span><span class="sxs-lookup"><span data-stu-id="a7bae-143">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="a7bae-144">尝试插入外接程序时在收到提示的时候验证其身份。</span><span class="sxs-lookup"><span data-stu-id="a7bae-144">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="a7bae-145">确认是否已安装最新的 Office 更新程序，或下载 [Office 2013 更新程序](https://support.microsoft.com/kb/2986156/)。</span><span class="sxs-lookup"><span data-stu-id="a7bae-145">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>

## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="a7bae-146">无法显示外接程序对话框</span><span class="sxs-lookup"><span data-stu-id="a7bae-146">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="a7bae-p109">使用 Office 外接程序时，将要求用户允许显示对话框。用户选择“**允许**”，将出现以下错误消息：</span><span class="sxs-lookup"><span data-stu-id="a7bae-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="a7bae-p110">“浏览器中的安全设置阻止创建对话框。请尝试使用其他浏览器，或者配置浏览器，使地址栏中显示的 [URL] 和域处于同一安全区域。”</span><span class="sxs-lookup"><span data-stu-id="a7bae-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![对话框错误消息的屏幕截图](../images/dialog-prevented.png)

|<span data-ttu-id="a7bae-152">**受影响的浏览器**</span><span class="sxs-lookup"><span data-stu-id="a7bae-152">**Affected browsers**</span></span>|<span data-ttu-id="a7bae-153">**受影响的平台**</span><span class="sxs-lookup"><span data-stu-id="a7bae-153">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="a7bae-154">Internet Explorer、Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="a7bae-154">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="a7bae-155">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="a7bae-155">Office on the web</span></span>|

<span data-ttu-id="a7bae-p111">若要解决此问题，最终用户或管理员可以向 Internet Explorer 中的受信任站点列表添加外接程序的域。无论使用的是 Internet Explorer 还是 Microsoft Edge 浏览器，请使用相同过程。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a7bae-158">请勿将不信任的加载项的 URL 添加到受信任网站列表中。</span><span class="sxs-lookup"><span data-stu-id="a7bae-158">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="a7bae-159">要将 URL 添加到受信任站点的列表中，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="a7bae-159">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="a7bae-160">在“**控制面板**”中，转到“**Internet 选项**” > “**安全性**”。</span><span class="sxs-lookup"><span data-stu-id="a7bae-160">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="a7bae-161">选择“**受信任站点**”区域，并选择“**网站**”。</span><span class="sxs-lookup"><span data-stu-id="a7bae-161">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="a7bae-162">输入错误消息中显示的 URL，然后选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="a7bae-162">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="a7bae-p112">再次尝试使用外接程序。如果问题仍然存在，请验证其他安全区域的设置，并确保外接程序域与 Office 应用程序地址栏中显示的 URL 处于同一区域。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="a7bae-p113">在弹出模式中使用对话框 API 时，会出现此问题。若要避免出现此问题，请使用 [displayInFrame](/javascript/api/office/office.ui) 标记。这要求页面支持在 iframe 中进行显示。以下示例演示如何使用此标记。</span><span class="sxs-lookup"><span data-stu-id="a7bae-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="see-also"></a><span data-ttu-id="a7bae-169">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a7bae-169">See also</span></span>

- [<span data-ttu-id="a7bae-170">Office 加载项开发错误疑难解答</span><span class="sxs-lookup"><span data-stu-id="a7bae-170">Troubleshoot development errors with Office Add-ins</span></span>](troubleshoot-development-errors.md)
