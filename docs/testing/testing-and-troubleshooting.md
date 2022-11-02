---
title: 排查 Office 加载项中的用户错误
description: 了解如何排查 Office 加载项中的用户错误。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 18bb3c180cd3af1eb8d045d7c69b9772532b04d4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810370"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>排查 Office 加载项中的用户错误

At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.

还可以使用 [Fiddler](https://www.telerik.com/fiddler) 发现和调试加载项问题。

## <a name="common-errors-and-troubleshooting-steps"></a>常见错误和故障排除步骤

下表列出了用户可能遇到的常见错误消息以及用户可以采取以解决这些错误的步骤。

|**错误消息**|**解决方案**|
|:-----|:-----|
|应用错误：无法访问目录|Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.|
|APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.|确认已安装最新的 Office 更新，或下载 [Office 2013 更新](https://support.microsoft.com/kb/2986156/)。|
|错误：对象不 支持此属性或方法 "defineProperty"|确认 Internet Explorer 不是在兼容模式下运行。 转到 **“工具** > **兼容性视图设置”。**|
|很抱歉，我们无法加载 该应用程序，因为您的浏览器 版本不受支持。 单击此处查看 支持的浏览器版本的列表。|Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a>安装加载项时，状态栏中会显示“加载加载项时出错”

1. 关闭 Office。
1. 验证清单是否有效。 请参阅 [验证 Office 外接程序清单](troubleshoot-manifest.md)。
1. 重新启动外接程序。
1. 再次安装加载项。

你还可以向我们提供反馈：如果使用 Windows 版 Excel 或 Mac 版 Excel，可以直接从 Excel 向 Office 扩展性团队报告反馈。 若要执行此操作，请选择“**文件**” > “**反馈**” > “**发送哭脸**”。 发送哭脸将提供必要的日志，以帮助我们了解该问题。

## <a name="outlook-add-in-doesnt-work-correctly"></a>Outlook 外接程序不能正常工作

如果在 Windows 上运行并[使用 Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) 的 Outlook 加载项不能正常工作，请尝试在 Internet Explorer 中启用脚本调试。

- 转到 **“工具** > **”“Internet 选项** > **高级**”。
- 在“浏览”下，取消选中“禁用脚本调试 (Internet Explorer)”和“禁用脚本调试 (其他)”。

We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.

## <a name="add-in-doesnt-activate-in-office-2013"></a>外接程序在 Office 2013 中无法激活

如果用户执行以下步骤时加载项未激活。

1. 使用 Microsoft 帐户在 Office 2013 中登录。

1. 为其 Microsoft 帐户启用两步验证。

1. 尝试插入外接程序时在收到提示的时候验证其身份。

确认是否已安装最新的 Office 更新程序，或下载 [Office 2013 更新程序](https://support.microsoft.com/kb/2986156/)。

## <a name="add-in-dialog-box-cannot-be-displayed"></a>无法显示外接程序对话框

使用 Office 外接程序时，将要求用户允许显示对话框。 用户选择 **“允许”**，并出现以下错误消息。

"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."

![对话框错误消息的屏幕截图。](../images/dialog-prevented.png)

|受影响的浏览器|受影响的平台|
|:--------------------|:---------------------|
|Microsoft Edge|Office 网页版|

若要解决此问题，最终用户或管理员可以将外接程序的域添加到 Microsoft Edge 浏览器中的受信任站点列表。

> [!IMPORTANT]
> 请勿将不信任的加载项的 URL 添加到受信任网站列表中。

要将 URL 添加到受信任站点的列表中，请执行以下操作：

1. 在“**控制面板**”中，转到“**Internet 选项**” > “**安全性**”。
1. 选择“**受信任站点**”区域，并选择“**网站**”。
1. 输入错误消息中显示的 URL，然后选择“**添加**”。
1. Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.

This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="see-also"></a>另请参阅

- [排查 Office 加载项中的开发错误](troubleshoot-development-errors.md)
