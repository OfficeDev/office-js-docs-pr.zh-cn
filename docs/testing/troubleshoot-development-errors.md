---
title: 排查 Office 加载项中的开发错误
description: 了解如何排查Office加载项中的开发错误。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7f463b7a7c9a8895283b9f8e18c11bdb63d3da9d
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091123"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a>排查 Office 加载项中的开发错误

下面是开发Office外接程序时可能会遇到的常见问题的列表。

> [!TIP]
> 清除Office缓存通常会解决与过时代码相关的问题。 这可保证使用当前文件名、菜单文本和其他命令元素上传最新的清单。 若要了解详细信息，请参阅[“清除Office缓存](clear-cache.md)。

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>外接程序无法在任务窗格中加载，或外接程序清单存在其他问题

请参阅[验证 Office 加载项的清单](troubleshoot-manifest.md)和[使用运行时日志记录功能调试加载项](runtime-logging.md)，以调试加载项清单问题。

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a>对加载项命令（包括功能区按钮和菜单项）的更改未生效

清除缓存有助于确保正在使用加载项清单的最新版本。 若要清除Office缓存，请按照[“清除Office缓存”中的](clear-cache.md)说明操作。 如果使用Office web 版，请通过浏览器的 UI 清除浏览器的缓存。

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a>对静态文件（例如 JavaScript、HTML 和 CSS）的更改未生效

浏览器可能正在缓存这些文件。 若要阻止此操作，请在开发时关闭客户端缓存。 详细信息取决于你所使用的服务器类型。 在大多数情况下，它涉及将某些标头添加到 HTTP 响应。 建议使用以下集。

- Cache-Control：“private、no-cache、no-store”
- Pragma：“no-cache”
- 过期：“-1”

有关在 Node.JS Express 服务器中执行此操作的示例，请参阅[此 app.js 文件](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/app.js)。 有关 ASP.NET 项目中的示例，请参阅[此 cshtml 文件](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml)。

如果加载项托管在 Internet Information Server (IIS) 中，则还可以将以下内容添加到 web.config 中。

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

如果这些步骤一开始似乎不起作用，则可能需要清除浏览器的缓存。 请通过浏览器的 UI 执行此操作。 有时，当你尝试在边缘 UI 中清除边缘缓存时，无法成功清除它。 如果出现这种情况，请在 Windows 命令提示符中运行以下命令。

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a>对属性值所做的更改不会发生，也没有错误消息

检查属性的参考文档，查看该属性是否为只读。 此外，Office JS [的 TypeScript 定义](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)指定哪些对象属性是只读的。 如果尝试设置只读属性，则写入操作将以无提示方式失败，不会引发任何错误。 以下示例错误地尝试设置只读属性 [Chart.id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member)。另请参阅 [无法直接设置某些属性](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly)。

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="getting-error-this-add-in-is-no-longer-available"></a>收到错误：“此加载项不再可用”

以下是导致此错误的一些原因。 如果发现其他原因，请使用页面底部的反馈工具告诉我们。

- 如果使用的是Visual Studio，则旁加载可能会出现问题。 关闭Office主机和Visual Studio的所有实例。 重启Visual Studio，然后再次尝试按 F5。
- 外接程序的清单已从其部署位置（例如集中部署、SharePoint目录或网络共享）中删除。
- 清单中 [ID](/javascript/api/manifest/id) 元素的值已直接在已部署的副本中更改。 如果出于任何原因，需要更改此 ID，请先从Office主机中删除加载项，然后将原始清单替换为已更改的清单。 许多人需要清除Office缓存才能删除原始文件的所有跟踪。 有关清除操作系统的缓存的说明，请参阅“[清除Office缓存](clear-cache.md)”一文。
- 外接程序的清单在`resid`清单的[“资源](/javascript/api/manifest/resources)”部分的任何位置都没有定义，或者在使用该清单的位置和在节中`<Resources>`定义的位置之间的拼写`resid`不匹配。
- 清单中有一个 `resid` 超过 32 个字符的属性。 属性 `resid` 和 `id` 节中 `<Resources>` 相应资源的属性不能超过 32 个字符。
- 外接程序具有自定义加载项命令，但你尝试在不支持加载项的平台上运行它。 有关详细信息，请参阅 [加载项命令要求集](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)。

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a>加载项在 Edge 上不起作用，但它适用于其他浏览器

请参阅[Microsoft Edge问题疑难解答](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)。

## <a name="excel-add-in-throws-errors-but-not-consistently"></a>Excel加载项引发错误，但不一致

有关可能的原因，请参阅[Excel加载项疑难解答](../excel/excel-add-ins-troubleshooting.md)。

## <a name="manifest-schema-validation-errors-in-visual-studio-projects"></a>Visual Studio项目中的清单架构验证错误

如果使用的是需要更改清单文件的较新功能，则可能会在Visual Studio中收到验证错误。 例如，添加 `<Runtimes>` 元素以实现共享 JavaScript 运行时时，可能会看到以下验证错误。

**命名空间“”中的元素“http://schemas.microsoft.com/office/taskpaneappversionoverridesHost”在命名空间http://schemas.microsoft.com/office/taskpaneappversionoverrides“”中具有无效的子元素“Runtimes”**

如果发生这种情况，可以将Visual Studio使用的 XSD 文件更新到最新版本。 最新的架构版本为 [[MS-OWEMXML]： 附录 A： 完整 XML 架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)。

### <a name="locate-the-xsd-files"></a>找到 XSD 文件

1. 在 Visual Studio 中打开项目。
1. 在 **解决方案资源管理器** 中，打开manifest.xml文件。 清单通常位于解决方案下的第一个项目中。
1. 选择“ **查看** > **属性”窗口** (F4) 。
1. 在 **“属性”窗口** 中，选择省略号 (...) 打开 **XML 架构编辑器** 。 可在此处找到项目使用的所有架构文件的确切文件夹位置。

### <a name="update-the-xsd-files"></a>更新 XSD 文件

1. 在文本编辑器中打开要更新的 XSD 文件。 验证错误中的架构名称将与 XSD 文件名相关联。 例如，打开 **TaskPaneAppVersionOverridesV1_0.xsd**。
1. 在 [[MS-OWEMXML]： 附录 A： 完整 XML 架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)处找到更新后的架构。 例如，TaskPaneAppVersionOverridesV1_0位于 [taskpaneappversionoverrides 架构](/openspecs/office_file_formats/ms-owemxml/82e93ec5-de22-42a8-86e3-353c8336aa40)中。
1. 将文本复制到文本编辑器中。
1. 保存更新后的 XSD 文件。
1. 重启Visual Studio以选取新的 XSD 文件更改。

对于过期的任何其他架构，可以重复上一个过程。

## <a name="when-working-offline-no-office-apis-work"></a>脱机工作时，Office API 不起作用

从本地副本而不是从CDN加载 Office JavaScript 库时，如果该库不是最新的，则 API 可能会停止工作。 如果已离开项目一段时间，请重新安装库以获取最新版本。 此过程因 IDE 而异。 根据环境选择以下选项之一。

- **Visual Studio**：请参阅 [最新 Office JavaScript API 库的更新](../develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。 
- **任何其他 IDE**：请参阅npm包 [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) 和 [@types/office-js](https://www.npmjs.com/package/@types/office-js)。

## <a name="see-also"></a>另请参阅

- [在 Office 网页版中调试加载项](debug-add-ins-in-office-online.md)
- [将 Office 外接程序旁加载到 iPad 和 Mac 上](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [在 Mac 上调试 Office 加载项](debug-office-add-ins-on-ipad-and-mac.md)  
- [适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展](debug-with-vs-extension.md)
- [验证 Office 加载项的清单](troubleshoot-manifest.md)
- [使用运行时日志记录功能调试加载项](runtime-logging.md)
- [排查 Office 加载项中的用户错误](testing-and-troubleshooting.md)
- [Microsoft Q&A (office-js-dev) ](/answers/topics/office-js-dev.html)
