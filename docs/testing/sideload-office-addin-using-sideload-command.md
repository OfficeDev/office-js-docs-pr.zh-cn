---
title: 使用 sideload 命令旁加载 Office 外接程序
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: e831a1dfbc31ecf06c8b2d78dc1e9a8a4c9dcf01
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/27/2018
ms.locfileid: "21279358"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>旁加载 Office 外接程序，以供使用 **sideload 命令**测试
 >[!NOTE]
>"npm run sideload" 方法仅适用于 Excel、Word 和 PowerPoint 外接程序）。

1. 以管理员身份打开命令提示符：

2. 将目录更改为外接程序项目文件夹的根目录。

3. 运行以下命令以，在端口 3000 上启动本地 Web 服务器实例以提供外接程序项目："**npm run start**"

4. 以管理员身份打开第二个命令提示符。

5. 将目录更改为外接程序项目文件夹的根目录。

6. 运行以下命令以引导主机应用程序（例如 Excel、Word）并在主机应用程序中注册您的外接程序："**npm run sideload**"

## <a name="see-also"></a>另请参阅

- [验证并排查清单问题](troubleshoot-manifest.md)
- [发布 Office 外接程序](../publish/publish.md)