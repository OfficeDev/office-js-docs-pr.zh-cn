---
title: 清除 Office 缓存
description: 了解如何清除计算机上的 Office 缓存。
ms.date: 11/15/2021
ms.localizationpriority: high
ms.openlocfilehash: 79b5f4e483eadec5d9f3095ab1c37e8eb697658b
ms.sourcegitcommit: 6e6c4803fdc0a3cc2c1bcd275288485a987551ff
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2021
ms.locfileid: "61066665"
---
# <a name="clear-the-office-cache"></a>清除 Office 缓存

你需要清除计算机上的 Office 缓存来删除以前在 Windows、Mac 或 iOS 上旁加载的加载项。

此外，如果你对加载项的清单进行了更改（例如，更新图标的文件名或加载项命令的文本），则应清除 Office 缓存，然后使用更新后的清单重新旁加载此加载项。执行此操作后，Office 将按照更新清单中所述的方式呈现该加载项。

> [!NOTE]
> 若要从 Excel、OneNote、PowerPoint 或 Word 网页版中删除旁加载的加载项，请参阅 [在 Office 网页版中旁加载 Office 加载项以进行测试：删除旁加载的加载项](sideload-office-add-ins-for-testing.md#remove-a-sideloaded-add-in)。

## <a name="clear-the-office-cache-on-windows"></a>清除 Windows 上的 Office 缓存

在 Windows 计算机上清除 Office 缓存的方法有三种：自动、手动以及使用 Microsoft Edge 开发人员工具。 以下子部分介绍了这些方法。

### <a name="automatically"></a>自动

建议对加载项开发计算机使用此方法。 如果 Windows 版 Office 的版本为 2108 或更高版本，则以下步骤会将 Office 缓存配置为每次重新打开 Office 时自动清除。

> [!NOTE]
> Outlook 不支持自动方法。

1. 从除 Outlook 以外的任何 Office 主机的功能区中，导航到“**文件**” > “**选项**” > “**信任中心**” > “**信任中心设置**” > “**受信任的加载项目录**”。
1. 选中“**下次启动 Office 时，清除以前启动的所有 Web 加载项的缓存**”复选框。

### <a name="manually"></a>手动

Excel、Word 和 PowerPoint 的手动方法与 Outlook 不同。

#### <a name="manually-clear-the-cache-in-excel-word-and-powerpoint"></a>手动清除 Excel、Word 和 PowerPoint 中的缓存

若要从 Excel、Word 和 PowerPoint 中删除所有旁加载的加载项，请删除以下文件夹中的内容。

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

如果存在以下文件夹，则也删除其内容：

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

#### <a name="manually-clear-the-cache-in-outlook"></a>手动清除 Outlook 中的缓存

要从 Outlook 中删除旁加载的加载项，请使用 [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md) 中概述的步骤，以在列出已安装加载项的对话框的“**自定义加载项**”部分中查找该加载项。选择加载项所对应的省略号（`...`），然后选择“**删除**”以删除该特定加载项。如果加载项删除不起作用，则删除 `Wef` 的内容，如之前 针对 Excel、Word 和 PowerPoint 所述。

### <a name="using-the-microsoft-edge-developer-tools"></a>使用 Microsoft Edge 开发人员工具

若要在加载项在 Microsoft Edge 中运行时清除 Windows 10 上的 Office 缓存，可以使用 Microsoft Edge DevTools。

> [!TIP]
> 如果只希望旁加载的加载项反映对其 HTML 或 JavaScript 源文件的最新更改，则应该不需要清除缓存。 相反，只需将焦点放在加载项的任务窗格中（通过单击任务窗格中的任意位置），然后按 **Ctrl+F5** 重新加载加载项。

> [!NOTE]
> 要使用以下步骤清除 Office 缓存，加载项必须具有任务窗格。如果加载项是无 UI 的加载项（例如，使用 [on-send](../outlook/outlook-on-send-addins.md) 功能的加载项），将需要先为加载项添加一个任务窗格，且该任务窗格使用与 [SourceLocation](../reference/manifest/sourcelocation.md) 相同的域，然后才能使用以下步骤来清除缓存。

1. 安装 [Microsoft Edge 开发工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj)。

2. 在 Office 客户端中打开加载项。

3. 运行 Microsoft Edge 开发工具。

4. 在 Microsoft Edge 开发工具中，打开“**本地**”选项卡。加载项将按其名称列出。

5. 选择加载项名称以将调试器连接到加载项。 当调试器连接到加载项时，将打开一个新的“Microsoft Edge 开发工具”窗口。

6. 在新窗口的“**网络**”选项卡上，选择“**清除缓存**”。

    ![Microsoft Edge 开发工具屏幕截图，其中突出显示了“清除缓存”按钮。](../images/edge-devtools-clear-cache.png)

7. 如果完成这些步骤后未获得想要的结果，请尝试选择“**始终从服务器中刷新**”。

    ![Microsoft Edge 开发工具屏幕截图，其中突出显示了“始终从服务器中刷新”按钮。](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a>清除 Mac 上的 Office 缓存

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a>清除 iOS 上的 Office 缓存

要清除 iOS 上的 Office 缓存，请从加载项中的 JavaScript 调用 `window.location.reload(true)` 以强制重新加载。或者重新安装 Office。

## <a name="see-also"></a>另请参阅

- [排查 Office 加载项中的开发错误](troubleshoot-development-errors.md)
- [使用适用于 Internet Explorer 的开发人员工具调试加载项](debug-add-ins-using-f12-tools-ie.md)
- [使用旧版 Edge 开发人员工具调试加载项](debug-add-ins-using-devtools-edge-legacy.md)
- [使用运行时日志记录功能调试加载项](runtime-logging.md)
- [旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [Office 加载项 XML 清单](../develop/add-in-manifests.md)
- [验证 Office 加载项的清单](troubleshoot-manifest.md)
