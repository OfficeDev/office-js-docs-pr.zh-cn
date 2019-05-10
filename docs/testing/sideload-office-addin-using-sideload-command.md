---
title: 使用旁加载命令旁加载 Office 加载项
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: 69d39c2736312653b5a362aefccd41629e6e3555
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619075"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>使用旁加载命令旁加载 Office 加载项以供测试
 
> [!NOTE]
> 本文中所述的旁加载技术仅适用于：
> 
> - Windows 上运行的 Excel、Word 和 PowerPoint 加载项
> 
> - 使用[适合于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建并且 package.json 文件中的 `scripts` 部分具有 `sideload` 脚本的加载项项目。 （使用更早版本的适用于 Office 加载项的 Yeoman 生成器创建的项目没有此脚本。）
 
若要使用适合于 Office 加载项的 Yeoman 生成器提供的 `sideload` 脚本旁加载你的加载项，请按照以下步骤操作：

1. 以管理员身份打开命令提示符。

2. 将目录更改为加载项项目文件夹的根目录。

3. 运行以下命令以在端口 3000 上启动本地 Web 服务器实例，以便为加载项项目提供服务：`npm run start`

4. 以管理员身份打开第二个命令提示符。

5. 将目录更改为加载项项目文件夹的根目录。

6. 运行以下命令以引导主机应用程序（例如 Excel、Word）并在主机应用程序中注册你的加载项：`npm run sideload`

如果你的加载项项目是使用 Visual Studio 创建的，或者没有旁加载脚本，则可以使用[从网络共享旁加载 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中所述的方法在 Windows 上旁加载它。

如果不在 Windows 上测试 Word、Excel 或 PowerPoint 加载项，则请参阅以下主题之一，以了解与旁加载你的加载项相关的信息：
 
- [在 Office Online 中旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [在 iPad 和 Mac 上旁加载 Office 外接程序进行测试](sideload-an-office-add-in-on-ipad-and-mac.md)
- [旁加载 Outlook 加载项以供测试](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a>另请参阅

- [验证并排查清单问题](troubleshoot-manifest.md)
- [发布 Office 外接程序](../publish/publish.md)
