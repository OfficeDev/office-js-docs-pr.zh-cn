---
title: 排查Office加载项的开发错误
description: 了解如何解决加载项中的Office错误。
ms.date: 06/11/2021
localization_priority: Normal
ms.openlocfilehash: 8f0ceaf13041fa27c4e9e279646e979f132913b3
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349273"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a><span data-ttu-id="a7f25-103">排查Office加载项的开发错误</span><span class="sxs-lookup"><span data-stu-id="a7f25-103">Troubleshoot development errors with Office Add-ins</span></span>

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="a7f25-104">外接程序无法在任务窗格中加载，或外接程序清单存在其他问题</span><span class="sxs-lookup"><span data-stu-id="a7f25-104">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="a7f25-105">请参阅[验证 Office 加载项的清单](troubleshoot-manifest.md)和[使用运行时日志记录功能调试加载项](runtime-logging.md)，以调试加载项清单问题。</span><span class="sxs-lookup"><span data-stu-id="a7f25-105">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="a7f25-106">对加载项命令（包括功能区按钮和菜单项）的更改未生效</span><span class="sxs-lookup"><span data-stu-id="a7f25-106">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="a7f25-107">如果在清单中进行的更改（如功能区按钮图标的文件名或菜单项的文本）似乎没有生效，请尝试清除计算机上的 Office 缓存。</span><span class="sxs-lookup"><span data-stu-id="a7f25-107">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="a7f25-108">对于 Windows：</span><span class="sxs-lookup"><span data-stu-id="a7f25-108">For Windows:</span></span>

<span data-ttu-id="a7f25-109">删除文件夹的内容， `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 并删除文件夹的内容 `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` （如果存在）。</span><span class="sxs-lookup"><span data-stu-id="a7f25-109">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="a7f25-110">对于 Mac：</span><span class="sxs-lookup"><span data-stu-id="a7f25-110">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="a7f25-111">对于 iOS：</span><span class="sxs-lookup"><span data-stu-id="a7f25-111">For iOS:</span></span>

<span data-ttu-id="a7f25-p101">在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="a7f25-p101">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="a7f25-114">对静态文件（例如 JavaScript、HTML 和 CSS）的更改未生效</span><span class="sxs-lookup"><span data-stu-id="a7f25-114">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="a7f25-115">浏览器可能正在缓存这些文件。</span><span class="sxs-lookup"><span data-stu-id="a7f25-115">The browser may be caching these files.</span></span> <span data-ttu-id="a7f25-116">若要阻止此操作，请在开发时关闭客户端缓存。</span><span class="sxs-lookup"><span data-stu-id="a7f25-116">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="a7f25-117">详细信息取决于你所使用的服务器类型。</span><span class="sxs-lookup"><span data-stu-id="a7f25-117">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="a7f25-118">在大多数情况下，它涉及将某些标头添加到 HTTP 响应。</span><span class="sxs-lookup"><span data-stu-id="a7f25-118">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="a7f25-119">建议设置以下集合。</span><span class="sxs-lookup"><span data-stu-id="a7f25-119">We suggest the following set.</span></span>

- <span data-ttu-id="a7f25-120">Cache-Control：“private、no-cache、no-store”</span><span class="sxs-lookup"><span data-stu-id="a7f25-120">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="a7f25-121">Pragma：“no-cache”</span><span class="sxs-lookup"><span data-stu-id="a7f25-121">Pragma: "no-cache"</span></span>
- <span data-ttu-id="a7f25-122">过期：“-1”</span><span class="sxs-lookup"><span data-stu-id="a7f25-122">Expires: "-1"</span></span>

<span data-ttu-id="a7f25-123">有关在 Node.JS Express 服务器中执行此操作的示例，请参阅[此 app.js 文件](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)。</span><span class="sxs-lookup"><span data-stu-id="a7f25-123">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="a7f25-124">有关 ASP.NET 项目中的示例，请参阅[此 cshtml 文件](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)。</span><span class="sxs-lookup"><span data-stu-id="a7f25-124">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="a7f25-125">如果加载项托管在 Internet Information Server (IIS) 中，则还可以将以下内容添加到 web.config 中。</span><span class="sxs-lookup"><span data-stu-id="a7f25-125">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="a7f25-126">如果这些步骤一开始似乎不起作用，则可能需要清除浏览器的缓存。</span><span class="sxs-lookup"><span data-stu-id="a7f25-126">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="a7f25-127">请通过浏览器的 UI 执行此操作。</span><span class="sxs-lookup"><span data-stu-id="a7f25-127">Do this through the UI of the browser.</span></span> <span data-ttu-id="a7f25-128">有时，当你尝试在边缘 UI 中清除边缘缓存时，无法成功清除它。</span><span class="sxs-lookup"><span data-stu-id="a7f25-128">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="a7f25-129">如果出现这种情况，请在 Windows 命令提示符中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="a7f25-129">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a><span data-ttu-id="a7f25-130">对属性值所做的更改不会发生，并且没有错误消息</span><span class="sxs-lookup"><span data-stu-id="a7f25-130">Changes made to property values don't happen and there is no error message</span></span>

<span data-ttu-id="a7f25-131">查看属性的参考文档，以查看其是否只读。</span><span class="sxs-lookup"><span data-stu-id="a7f25-131">Check the reference documentation for the property to see if it is read only.</span></span> <span data-ttu-id="a7f25-132">此外[，JS 的 TypeScript](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) Office指定哪些对象属性是只读的。</span><span class="sxs-lookup"><span data-stu-id="a7f25-132">Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="a7f25-133">如果您尝试设置只读属性，写入操作将失败，无提示，不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="a7f25-133">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="a7f25-134">以下示例错误地尝试将只读属性设置为 [Chart.id](/javascript/api/excel/excel.chart#id)。另请参阅 [某些属性不能直接设置](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。</span><span class="sxs-lookup"><span data-stu-id="a7f25-134">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a><span data-ttu-id="a7f25-135">收到错误："此外接程序不再可用"</span><span class="sxs-lookup"><span data-stu-id="a7f25-135">Getting error: "This add-in is no longer available"</span></span>

<span data-ttu-id="a7f25-136">以下是导致此错误的一些原因。</span><span class="sxs-lookup"><span data-stu-id="a7f25-136">The following are some of the causes of this error.</span></span> <span data-ttu-id="a7f25-137">如果发现其他原因，请使用页面底部的反馈工具告诉我们。</span><span class="sxs-lookup"><span data-stu-id="a7f25-137">If you discover additional causes, please tell us with the feedback tool at the bottom of the page.</span></span>

- <span data-ttu-id="a7f25-138">如果使用 Visual Studio，则旁加载可能有问题。</span><span class="sxs-lookup"><span data-stu-id="a7f25-138">If you are using Visual Studio, there may be a problem with the sideloading.</span></span> <span data-ttu-id="a7f25-139">关闭主机和Office的所有Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="a7f25-139">Close all instances of the Office host and Visual Studio.</span></span> <span data-ttu-id="a7f25-140">重新启动Visual Studio并再次尝试按 F5。</span><span class="sxs-lookup"><span data-stu-id="a7f25-140">Restart Visual Studio and try pressing F5 again.</span></span>
- <span data-ttu-id="a7f25-141">外接程序的清单已从其部署位置（如集中部署、SharePoint目录或网络共享）中删除。</span><span class="sxs-lookup"><span data-stu-id="a7f25-141">The add-in's manifest has been removed from its deployment location, such as Centralized Deployment, a SharePoint catalog, or a network share.</span></span>
- <span data-ttu-id="a7f25-142">清单 [中 ID](../reference/manifest/id.md) 元素的值已在已部署的副本中直接更改。</span><span class="sxs-lookup"><span data-stu-id="a7f25-142">The value of the [ID](../reference/manifest/id.md) element in the manifest has been changed directly in the deployed copy.</span></span> <span data-ttu-id="a7f25-143">如果出于任何原因需要更改此 ID，请首先从 Office 主机中删除外接程序，然后将原始清单替换为已更改的清单。</span><span class="sxs-lookup"><span data-stu-id="a7f25-143">If for any reason, you want to change this ID, first remove the add-in from the Office host, then replace the original manifest with the changed manifest.</span></span> <span data-ttu-id="a7f25-144">许多用户需要清除Office缓存以删除原始缓存的所有跟踪。</span><span class="sxs-lookup"><span data-stu-id="a7f25-144">You many need to clear the Office cache to remove all traces of the original.</span></span> <span data-ttu-id="a7f25-145">请参阅本文前面 [对外接程序命令（包括功能区按钮](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) 和菜单项）的更改不会生效一节。</span><span class="sxs-lookup"><span data-stu-id="a7f25-145">See the section [Changes to add-in commands including ribbon buttons and menu items do not take effect](#changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect) earlier in this article.</span></span>
- <span data-ttu-id="a7f25-146">外接程序的清单有 一个 未在清单的 Resources 部分的任何位置定义的 ，或者其使用位置和在 部分中定义位置之间的拼写不匹配。 `resid` [](../reference/manifest/resources.md) `resid` `<Resources>`</span><span class="sxs-lookup"><span data-stu-id="a7f25-146">The add-in's manifest has a `resid` that is not defined anywhere in the [Resources](../reference/manifest/resources.md) section of the manifest, or there is a mismatch in the spelling of the `resid` between where it is used and where it is defined in the `<Resources>` section.</span></span>
- <span data-ttu-id="a7f25-147">清单 `resid` 中的某位置有一个超过 32 个字符的属性。</span><span class="sxs-lookup"><span data-stu-id="a7f25-147">There is a `resid` attribute somewhere in the manifest with more than 32 characters.</span></span> <span data-ttu-id="a7f25-148">属性和节中相应资源的属性不能超过 `resid` `id` `<Resources>` 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="a7f25-148">A `resid` attribute, and the `id` attribute of the corresponding resource in the `<Resources>` section, cannot be more than 32 characters.</span></span>
- <span data-ttu-id="a7f25-149">加载项具有自定义加载项命令，但尝试在不支持命令的平台上运行。</span><span class="sxs-lookup"><span data-stu-id="a7f25-149">The add-in has a custom Add-in Command but you are trying to run it on a platform that doesn't support them.</span></span> <span data-ttu-id="a7f25-150">有关详细信息，请参阅加载项 [命令要求集](../reference/requirement-sets/add-in-commands-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="a7f25-150">For more information, see [Add-in commands requirement sets](../reference/requirement-sets/add-in-commands-requirement-sets.md).</span></span>

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a><span data-ttu-id="a7f25-151">外接程序在 Edge 上不起作用，但它适用于其他浏览器</span><span class="sxs-lookup"><span data-stu-id="a7f25-151">Add-in doesn't work on Edge but it works on other browsers</span></span>

<span data-ttu-id="a7f25-152">请参阅[疑难Microsoft Edge疑难解答](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。</span><span class="sxs-lookup"><span data-stu-id="a7f25-152">See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span>

## <a name="excel-add-in-throws-errors-but-not-consistently"></a><span data-ttu-id="a7f25-153">Excel加载项抛出错误，但不一致</span><span class="sxs-lookup"><span data-stu-id="a7f25-153">Excel add-in throws errors, but not consistently</span></span>

<span data-ttu-id="a7f25-154">请参阅[Excel加载项疑难](../excel/excel-add-ins-troubleshooting.md)解答了解可能的原因。</span><span class="sxs-lookup"><span data-stu-id="a7f25-154">See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.</span></span>

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a><span data-ttu-id="a7f25-155">清单架构验证错误Visual Studio项目中</span><span class="sxs-lookup"><span data-stu-id="a7f25-155">Manifest schema validation errors in Visual Studio projects</span></span>

<span data-ttu-id="a7f25-156">如果你使用的是需要更改清单文件的较新功能，你可能会在清单文件中收到Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="a7f25-156">If you are using newer features that require changes to the manifest file, you may get validation errors in Visual Studio.</span></span> <span data-ttu-id="a7f25-157">例如，添加 元素 `<Runtimes>` 来实现共享的 JavaScript 运行时时，你可能会看到以下验证错误。</span><span class="sxs-lookup"><span data-stu-id="a7f25-157">For example, when adding the `<Runtimes>` element to implement the shared JavaScript runtime, you may see the following validation error.</span></span>

<span data-ttu-id="a7f25-158">**命名空间 中的元素"Host"' 在命名空间 ' 中具有无效 http://schemas.microsoft.com/office/taskpaneappversionoverrides 的子元素 http://schemas.microsoft.com/office/taskpaneappversionoverrides "Runtimes"**</span><span class="sxs-lookup"><span data-stu-id="a7f25-158">**The element 'Host' in namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides' has invalid child element 'Runtimes' in namespace 'http://schemas.microsoft.com/office/taskpaneappversionoverrides'**</span></span>

<span data-ttu-id="a7f25-159">如果发生这种情况，你可以将 XSD 文件更新Visual Studio最新版本。</span><span class="sxs-lookup"><span data-stu-id="a7f25-159">If this occurs, you can update the XSD files that Visual Studio uses to the latest versions.</span></span> <span data-ttu-id="a7f25-160">最新架构版本位于 [[MS-OWEMXML]：附录 A：完全 XML 架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)。</span><span class="sxs-lookup"><span data-stu-id="a7f25-160">The latest schema versions are at [[MS-OWEMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span></span>

### <a name="locate-the-xsd-files"></a><span data-ttu-id="a7f25-161">找到 XSD 文件</span><span class="sxs-lookup"><span data-stu-id="a7f25-161">Locate the XSD files</span></span>

1. <span data-ttu-id="a7f25-162">在 Visual Studio 中打开项目。</span><span class="sxs-lookup"><span data-stu-id="a7f25-162">Open your project in Visual Studio.</span></span>
1. <span data-ttu-id="a7f25-163">在 **"解决方案资源管理器**"中，打开manifest.xml文件。</span><span class="sxs-lookup"><span data-stu-id="a7f25-163">In **Solution Explorer**, open the manifest.xml file.</span></span> <span data-ttu-id="a7f25-164">清单通常位于解决方案下的第一个项目中。</span><span class="sxs-lookup"><span data-stu-id="a7f25-164">The manifest is typically in the first project under your solution.</span></span>
1. <span data-ttu-id="a7f25-165">选择 **"查看**  >  **属性窗口**" (F4) 。</span><span class="sxs-lookup"><span data-stu-id="a7f25-165">Choose **View** > **Properties Window** (F4).</span></span>
1. <span data-ttu-id="a7f25-166">在" **属性窗口**"中，选择省略号" (...) "以打开 **XML 架构** 编辑器。</span><span class="sxs-lookup"><span data-stu-id="a7f25-166">In the **Properties Window**, choose the ellipsis (...) to open the **XML Schemas** editor.</span></span> <span data-ttu-id="a7f25-167">你可以在此处找到项目使用的所有架构文件的确切文件夹位置。</span><span class="sxs-lookup"><span data-stu-id="a7f25-167">Here you can find the exact folder location of all schema files your project uses.</span></span>

### <a name="update-the-xsd-files"></a><span data-ttu-id="a7f25-168">更新 XSD 文件</span><span class="sxs-lookup"><span data-stu-id="a7f25-168">Update the XSD files</span></span>

1. <span data-ttu-id="a7f25-169">在文本编辑器中打开要更新的 XSD 文件。</span><span class="sxs-lookup"><span data-stu-id="a7f25-169">Open the XSD file you want to update in a text editor.</span></span> <span data-ttu-id="a7f25-170">验证错误中的架构名称将关联到 XSD 文件名。</span><span class="sxs-lookup"><span data-stu-id="a7f25-170">The schema name from the validation error will correlate to the XSD file name.</span></span> <span data-ttu-id="a7f25-171">例如，打开 **TaskPaneAppVersionOverridesV1_0.xsd**。</span><span class="sxs-lookup"><span data-stu-id="a7f25-171">For example, open **TaskPaneAppVersionOverridesV1_0.xsd**.</span></span>
1. <span data-ttu-id="a7f25-172">找到更新后的架构 [，位置为 [MS-一级：XML 架构]：附录 A：完整 XML 架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)。</span><span class="sxs-lookup"><span data-stu-id="a7f25-172">Locate the updated schema at [[MS-OWEMXML]: Appendix A: Full XML Schema](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span></span> <span data-ttu-id="a7f25-173">例如，TaskPaneAppVersionOverridesV1_0位于 [taskpaneappversionoverrides Schema](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)中。</span><span class="sxs-lookup"><span data-stu-id="a7f25-173">For example, TaskPaneAppVersionOverridesV1_0 is at [taskpaneappversionoverrides Schema](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40).</span></span>
1. <span data-ttu-id="a7f25-174">将文本复制到文本编辑器中。</span><span class="sxs-lookup"><span data-stu-id="a7f25-174">Copy the text into your text editor.</span></span>
1. <span data-ttu-id="a7f25-175">保存更新后的 XSD 文件。</span><span class="sxs-lookup"><span data-stu-id="a7f25-175">Save the updated XSD file.</span></span>
1. <span data-ttu-id="a7f25-176">重新启动Visual Studio以选取新的 XSD 文件更改。</span><span class="sxs-lookup"><span data-stu-id="a7f25-176">Restart Visual Studio to pick up the new XSD file changes.</span></span>

<span data-ttu-id="a7f25-177">您可以对过期的其他任何架构重复上述过程。</span><span class="sxs-lookup"><span data-stu-id="a7f25-177">You can repeat the previous process for any additional schemas that are out-of-date.</span></span>

## <a name="see-also"></a><span data-ttu-id="a7f25-178">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a7f25-178">See also</span></span>

- [<span data-ttu-id="a7f25-179">在 Office 网页版中调试加载项</span><span class="sxs-lookup"><span data-stu-id="a7f25-179">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="a7f25-180">将 Office 外接程序旁加载到 iPad 和 Mac 上</span><span class="sxs-lookup"><span data-stu-id="a7f25-180">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="a7f25-181">在 iPad 和 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="a7f25-181">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="a7f25-182">适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展</span><span class="sxs-lookup"><span data-stu-id="a7f25-182">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="a7f25-183">验证 Office 加载项的清单</span><span class="sxs-lookup"><span data-stu-id="a7f25-183">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="a7f25-184">使用运行时日志记录功能调试加载项</span><span class="sxs-lookup"><span data-stu-id="a7f25-184">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="a7f25-185">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="a7f25-185">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
