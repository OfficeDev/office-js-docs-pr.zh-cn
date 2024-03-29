---
title: 旁加载 Office 加载项以从网络共享进行测试
description: 了解如何旁加载 Office 加载项以从网络共享进行测试。
ms.date: 05/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 87bdeb6cbd33bcd9b1828c7afa0a9f879d4c05e4
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712774"
---
# <a name="sideload-office-add-ins-for-testing-from-a-network-share"></a>旁加载 Office 加载项以从网络共享进行测试

可以通过将清单发布到网络文件共享，在 Windows 上的 Office 客户端中测试 Office 外接程序， (下面的说明) 。 在本地主机上完成开发和测试并想要从非本地服务器或云帐户测试外接程序时，应使用此部署选项。

> [!IMPORTANT]
> 生产加载项不支持按网络共享进行部署。此方法具有以下限制。
>
> - 加载项只能安装在 Windows 计算机上。
> - 如果加载项的新版本更改了功能区，例如向功能区添加自定义选项卡或自定义按钮，则每个用户都必须重新安装外接程序。

> [!NOTE]
> 如果你的外接程序项目是使用[外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md)的足够使用的版本，运行 `npm start` 时将自动在 Office 桌面客户端中旁加载外接程序。

本文仅适用于测试 Word、Excel、PowerPoint 和 Project 加载项，仅适用于 Windows。 如果要在另一个平台上进行测试或想要测试 Outlook 加载项，请参阅以下主题之一来旁加载外接程序。

- [在 Office 网页版中旁加载 Office 加载项进行测试](sideload-office-add-ins-for-testing.md)
- [在 Mac 上旁加载 Office 加载项以进行测试](sideload-an-office-add-in-on-mac.md)
- [在 iPad 上旁加载 Office 加载项以进行测试](sideload-an-office-add-in-on-ipad.md)
- [旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)

下面的视频逐步展示了如何使用共享文件夹目录在 Office 网页版或桌面上旁加载加载项。  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a>共享文件夹

1. 在想要托管外接程序的 Windows 计算机上，转到你想用作共享文件夹目录的文件夹的父文件夹或驱动器号。

1. 打开要用作共享文件夹目录的文件夹的上下文菜单（右键单击该文件夹），然后选择“**属性**”。

1. 在“**属性**”对话框窗口中，打开“**共享**”选项卡，然后选择“**共享**”按钮。

    ![“文件夹属性”对话框，其中突出显示了“共享”选项卡和“共享”按钮。](../images/sideload-windows-properties-dialog.png)

1. 在 **网络访问** 对话框窗口中，添加你自己以及要与其共享加载项的任何其他用户和/或组。 你至少需要对该文件夹的 **读/写** 权限。 选择要与其共享的人员后，请选择“**共享**”按钮。

1. 当你看到确认 **你的文件夹已共享** 的消息时，请记下紧跟文件夹名称显示的完整网络路径。 （当你 [将共享文件夹指定为受信任的目录](#specify-the-shared-folder-as-a-trusted-catalog)时，你需要将此值输入为 **目录UR**，如本文下一节所述。）选择“**完成**”按钮以关闭“**网络访问**”对话框窗口。

   ![突出显示了共享路径的网络访问对话框。](../images/sideload-windows-network-access-dialog.png)

1. 选择“**关闭**”按钮以关闭“**属性**”对话框窗口。

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>将共享文件夹指定为受信任的目录

### <a name="configure-the-trust-manually"></a>手动配置信任

1. 在 Excel、Word、PowerPoint 或 Project 中打开一个新的文档。

1. 选择“文件”选项卡，然后选择“选项”。

1. 选择“**信任中心**”，然后选择“**信任中心设置**”按钮。

1. 选择“**受信任的加载项目录**”。

1. 在“**目录 Url**”框中，输入你之前 [共享](#share-a-folder)的文件夹的完整网络路径。 如果在共享文件夹时未能记下文件夹的完整网络路径，则可以从文件夹的“**属性**”对话框窗口中获取它，如以下屏幕截图所示。

    ![“文件夹属性”对话框，其中突出显示了“共享”选项卡和网络路径。](../images/sideload-windows-properties-dialog-2.png)

1. 在“**目录 Url**”框中输入文件夹的完整网络路径后，选择“**添加目录**”按钮。

1. 选中新添加项目的“**在菜单中显示**”复选框，然后选择“**确定**”按钮以关闭“**信任中心**”对话框窗口。 

    ![“信任中心”对话框，其中选择了目录。](../images/sideload-windows-trust-center-dialog.png)

1. 选择 **“确定** ”按钮以关闭 **“选项** ”对话框窗口。

1. 关闭并重新打开 Office 应用程序，以使更改生效。

### <a name="configure-the-trust-with-a-registry-script"></a>使用注册表脚本配置信任

1. 在文本编辑器中，创建名为 TrustNetworkShareCatalog.reg 的文件。

1. 将以下内容添加到文件。

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```

1. 在众多在线 GUID 生成工具中选用一个（例如 [GUID 生成器](https://guidgenerator.com/)）来生成一个随机 GUID，并在 TrustNetworkShareCatalog.reg 文件中，将 *两个位置* 的“-random-GUID-here-”字符串都替换为 GUID。 （应保留右侧 `{}` 符号）。

1. 将 `Url` 值替换为你之前[共享](#share-a-folder)的文件夹的完整网络路径。 （请注意，URL 中的所有 `\` 字符都必须成双出现。）如果在共享文件夹时未能记下文件夹的完整网络路径，则可从文件夹的“**属性**”对话框窗口中获取它，如以下屏幕截图所示。

    ![“文件夹属性”对话框，其中突出显示了“共享”选项卡和网络路径。](../images/sideload-windows-properties-dialog-2.png)

1. 文件现应如下所示。 将其保存。

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

1. 关闭 *所有* Office 应用程序。

1. 如同对任何可执行文件操作一样运行 TrustNetworkShareCatalog.reg，例如双击它。

## <a name="sideload-your-add-in"></a>旁加载加载项

1. 放入在共享文件夹目录中进行测试的所有加载项的清单 XML 文件。 请务必将 Web 应用程序本身部署到 Web 服务器。 请务必在清单文件的元素中 **\<SourceLocation\>** 指定 URL。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

    > [!NOTE]
    > 对于 Visual Studio 项目，请使用项目在文件夹中生成的 `{projectfolder}\bin\Debug\OfficeAppManifests` 清单。

1. 在 Excel、Word 或 PowerPoint 中，选择功能区上“**插入**”选项卡中的“**我的加载项**”。 在 Project 中，选择功能区“**Project**”选项卡上的“**我的加载项**”。

1. 在“**Office 外接程序**”对话框的顶部，选择“**共享文件夹**”。

1. 选择加载项的名称，然后选择“**添加**”以插入加载项。

## <a name="remove-a-sideloaded-add-in"></a>删除旁加载的加载项

可以通过清除计算机上的 Office 缓存来删除以前旁加载的加载项。 有关如何清除 Windows 上的缓存的详细信息，请参阅“ [清除 Office 缓存](clear-cache.md#clear-the-office-cache-on-windows)”一文。

## <a name="see-also"></a>另请参阅

- [验证 Office 加载项的清单](troubleshoot-manifest.md)
- [清除 Office 缓存](clear-cache.md)
- [发布 Office 外接程序](../publish/publish.md)
