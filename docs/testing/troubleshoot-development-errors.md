---
title: 解决 Office 外接程序的开发错误
description: 了解如何解决 Office 外接程序中的开发错误。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5801146165446352ec806f6f832e9976f96467ac
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409386"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a><span data-ttu-id="0838e-103">解决 Office 外接程序的开发错误</span><span class="sxs-lookup"><span data-stu-id="0838e-103">Troubleshoot development errors with Office Add-ins</span></span>

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="0838e-104">外接程序无法在任务窗格中加载，或外接程序清单存在其他问题</span><span class="sxs-lookup"><span data-stu-id="0838e-104">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="0838e-105">请参阅[验证 Office 加载项的清单](troubleshoot-manifest.md)和[使用运行时日志记录功能调试加载项](runtime-logging.md)，以调试加载项清单问题。</span><span class="sxs-lookup"><span data-stu-id="0838e-105">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="0838e-106">对加载项命令（包括功能区按钮和菜单项）的更改未生效</span><span class="sxs-lookup"><span data-stu-id="0838e-106">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="0838e-107">如果在清单中进行的更改（如功能区按钮图标的文件名或菜单项的文本）似乎没有生效，请尝试清除计算机上的 Office 缓存。</span><span class="sxs-lookup"><span data-stu-id="0838e-107">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="0838e-108">对于 Windows：</span><span class="sxs-lookup"><span data-stu-id="0838e-108">For Windows:</span></span>

<span data-ttu-id="0838e-109">删除该文件夹的内容 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` ，并删除该文件夹的内容 `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` （如果存在）。</span><span class="sxs-lookup"><span data-stu-id="0838e-109">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="0838e-110">对于 Mac：</span><span class="sxs-lookup"><span data-stu-id="0838e-110">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="0838e-111">对于 iOS：</span><span class="sxs-lookup"><span data-stu-id="0838e-111">For iOS:</span></span>
<span data-ttu-id="0838e-p101">在加载项中通过 JavaScript 调用 `window.location.reload(true)`，以强制重载。也可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="0838e-p101">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="0838e-114">对静态文件（例如 JavaScript、HTML 和 CSS）的更改未生效</span><span class="sxs-lookup"><span data-stu-id="0838e-114">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="0838e-115">浏览器可能正在缓存这些文件。</span><span class="sxs-lookup"><span data-stu-id="0838e-115">The browser may be caching these files.</span></span> <span data-ttu-id="0838e-116">若要阻止此操作，请在开发时关闭客户端缓存。</span><span class="sxs-lookup"><span data-stu-id="0838e-116">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="0838e-117">详细信息取决于你所使用的服务器类型。</span><span class="sxs-lookup"><span data-stu-id="0838e-117">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="0838e-118">在大多数情况下，它涉及将某些标头添加到 HTTP 响应。</span><span class="sxs-lookup"><span data-stu-id="0838e-118">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="0838e-119">我们建议使用以下设置：</span><span class="sxs-lookup"><span data-stu-id="0838e-119">We suggest the following set:</span></span>

- <span data-ttu-id="0838e-120">Cache-Control：“private、no-cache、no-store”</span><span class="sxs-lookup"><span data-stu-id="0838e-120">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="0838e-121">Pragma：“no-cache”</span><span class="sxs-lookup"><span data-stu-id="0838e-121">Pragma: "no-cache"</span></span>
- <span data-ttu-id="0838e-122">过期：“-1”</span><span class="sxs-lookup"><span data-stu-id="0838e-122">Expires: "-1"</span></span>

<span data-ttu-id="0838e-123">有关在 Node.JS Express 服务器中执行此操作的示例，请参阅[此 app.js 文件](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js)。</span><span class="sxs-lookup"><span data-stu-id="0838e-123">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="0838e-124">有关 ASP.NET 项目中的示例，请参阅[此 cshtml 文件](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)。</span><span class="sxs-lookup"><span data-stu-id="0838e-124">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="0838e-125">如果加载项托管在 Internet Information Server (IIS) 中，则还可以将以下内容添加到 web.config 中。</span><span class="sxs-lookup"><span data-stu-id="0838e-125">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="0838e-126">如果这些步骤一开始似乎不起作用，则可能需要清除浏览器的缓存。</span><span class="sxs-lookup"><span data-stu-id="0838e-126">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="0838e-127">请通过浏览器的 UI 执行此操作。</span><span class="sxs-lookup"><span data-stu-id="0838e-127">Do this through the UI of the browser.</span></span> <span data-ttu-id="0838e-128">有时，当你尝试在边缘 UI 中清除边缘缓存时，无法成功清除它。</span><span class="sxs-lookup"><span data-stu-id="0838e-128">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="0838e-129">如果出现这种情况，请在 Windows 命令提示符中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="0838e-129">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a><span data-ttu-id="0838e-130">对属性值所做的更改不会发生，也不会出现错误消息</span><span class="sxs-lookup"><span data-stu-id="0838e-130">Changes made to property values don't happen and there is no error message</span></span>

<span data-ttu-id="0838e-131">检查属性的参考文档，以查看该属性是否为只读。</span><span class="sxs-lookup"><span data-stu-id="0838e-131">Check the reference documentation for the property to see if it is read only.</span></span> <span data-ttu-id="0838e-132">此外，Office JS 的 [TypeScript 定义](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) 指定哪些对象属性是只读的。</span><span class="sxs-lookup"><span data-stu-id="0838e-132">Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="0838e-133">如果尝试设置只读属性，写入操作将无提示地失败，且不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="0838e-133">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="0838e-134">下面的示例错误地尝试设置只读属性 [Chart.id](/javascript/api/excel/excel.chart#id)。另请参阅 [一些属性不能直接设置](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。</span><span class="sxs-lookup"><span data-stu-id="0838e-134">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a><span data-ttu-id="0838e-135">外接不在边缘上，而是在其他浏览器上运行</span><span class="sxs-lookup"><span data-stu-id="0838e-135">Add-in doesn't work on Edge but it works on other browsers</span></span>

<span data-ttu-id="0838e-136">请参阅 [Microsoft Edge 问题故障排除](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。</span><span class="sxs-lookup"><span data-stu-id="0838e-136">See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span>

## <a name="excel-add-in-throws-errors-but-not-consistently"></a><span data-ttu-id="0838e-137">Excel 加载项引发错误，但不一致</span><span class="sxs-lookup"><span data-stu-id="0838e-137">Excel add-in throws errors, but not consistently</span></span>

<span data-ttu-id="0838e-138">有关可能的原因，请参阅 [Excel 加载项疑难解答](../excel/excel-add-ins-troubleshooting.md) 。</span><span class="sxs-lookup"><span data-stu-id="0838e-138">See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.</span></span>

## <a name="see-also"></a><span data-ttu-id="0838e-139">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0838e-139">See also</span></span>

- [<span data-ttu-id="0838e-140">在 Office 网页版中调试加载项</span><span class="sxs-lookup"><span data-stu-id="0838e-140">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="0838e-141">将 Office 外接程序旁加载到 iPad 和 Mac 上</span><span class="sxs-lookup"><span data-stu-id="0838e-141">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="0838e-142">在 iPad 和 Mac 上调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="0838e-142">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="0838e-143">适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展</span><span class="sxs-lookup"><span data-stu-id="0838e-143">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="0838e-144">验证 Office 加载项的清单</span><span class="sxs-lookup"><span data-stu-id="0838e-144">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="0838e-145">使用运行时日志记录功能调试加载项</span><span class="sxs-lookup"><span data-stu-id="0838e-145">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="0838e-146">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="0838e-146">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
