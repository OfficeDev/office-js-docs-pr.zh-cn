---
title: 使用 sideload 命令旁加载 Office 外接程序
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 1ab0277493f2899adb479c2f24b1635a881af3cc
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944039"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>旁加载 Office 外接程序，以供使用 **sideload 命令**测试
 >[!NOTE]
>"npm run sideload" 方法仅适用于 Windows 平台上运行的 Excel、Word和 PowerPoint 加载项；并仅适用于使用 [**yo office** 工具](https://github.com/OfficeDev/generator-office) 创建的并在 `sideload`  `scripts`  package.json 文件中包含脚本的加载项项目（使用较早版本 **yo office** 创建的项目部不包含该脚本。）如果您使用 Visual Studio 创建项目或并不包含  sideload  脚本，您可以在  Windows  中使用 [通过网络共享刷入 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中说明的方法进行刷入。
>
> 如果不在 Windows 上测试 Word、Excel 或 PowerPoint 加载项，则请参阅以下主题之一来旁加载外接程序：
> 
> - [在 Office Online 中旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
> - [在 iPad 和 Mac 上旁加载 Office 外接程序进行测试](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [旁加载 Outlook 加载项以供测试](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. 以管理员身份打开命令提示。

2. 将目录更改为外接程序项目文件夹的根目录。

3. 运行以下命令以，在端口 3000 上启动本地 Web 服务器实例以提供外接程序项目："**npm run start**"

4. 以管理员身份打开第二个命令提示。

5. 将目录更改为外接程序项目文件夹的根目录。

6. 运行以下命令以引导主机应用程序（例如 Excel、Word）并在主机应用程序中注册您的外接程序："**npm run sideload**"

## <a name="see-also"></a>另请参阅

- [验证并排查清单问题](troubleshoot-manifest.md)
- [发布 Office 外接程序](../publish/publish.md)