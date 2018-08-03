---
title: 使用 sideload 命令旁加载 Office 外接程序
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: c3b53a70b5696e422653350de18d99be16d1d597
ms.sourcegitcommit: 0d4d78e275249f0d4b6a6cf807b42b79890c3023
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2018
ms.locfileid: "21773592"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>旁加载 Office 外接程序，以供使用 **sideload 命令**测试
 >[!NOTE]
> “npm run sideload”方法仅适用于在 Windows 上运行的 Excel Word 和 PowerPoint 加载项；并且仅适用于使用 [ ** yo office** 工具](https://github.com/OfficeDev/generator-office)创建并在 package.json 文件的 `sideload`   部分有 `scripts`    脚本的加载项项目。 （使用 **yo office** 旧版本创建的项目也没有这个脚本。）如果你的项目是用 Visual Studio 创建的，或者没有旁加载脚本，你可以用[从网络共享旁加载 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中描述的方法在 Windows 上旁加载它。
>
> 如果不在 Windows 上测试 Word、Excel 或 PowerPoint 加载项，请参阅以下主题之一来旁加载加载项：
> 
> - [在 Office Online 中旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
> - [在 iPad 和 Mac 上旁加载 Office 加载项以供测试](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [旁加载 Outlook 加载项以供测试](../../../../outlook/add-insSideload Outlook Add-ins for testing)

1. 以管理员身份打开命令提示符。

2. 将目录更改为外接程序项目文件夹的根目录。

3. 运行以下命令以，在端口 3000 上启动本地 Web 服务器实例以提供外接程序项目："**npm run start**"

4. 以管理员身份打开第二个命令提示符。

5. 将目录更改为外接程序项目文件夹的根目录。

6. 运行以下命令以引导主机应用程序（例如 Excel、Word）并在主机应用程序中注册您的外接程序："**npm run sideload**"

## <a name="see-also"></a>另请参阅

- [验证并排查清单问题](troubleshoot-manifest.md)
- [发布 Office 外接程序](../publish/publish.md)