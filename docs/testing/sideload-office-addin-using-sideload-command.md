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
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="800b3-102">旁加载 Office 外接程序，以供使用 **sideload 命令**测试</span><span class="sxs-lookup"><span data-stu-id="800b3-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="800b3-103">"npm run sideload" 方法仅适用于 Excel、Word 和 PowerPoint 外接程序）。</span><span class="sxs-lookup"><span data-stu-id="800b3-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

1. <span data-ttu-id="800b3-104">以管理员身份打开命令提示符：</span><span class="sxs-lookup"><span data-stu-id="800b3-104">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="800b3-105">将目录更改为外接程序项目文件夹的根目录。</span><span class="sxs-lookup"><span data-stu-id="800b3-105">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="800b3-106">运行以下命令以，在端口 3000 上启动本地 Web 服务器实例以提供外接程序项目："**npm run start**"</span><span class="sxs-lookup"><span data-stu-id="800b3-106">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="800b3-107">以管理员身份打开第二个命令提示符。</span><span class="sxs-lookup"><span data-stu-id="800b3-107">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="800b3-108">将目录更改为外接程序项目文件夹的根目录。</span><span class="sxs-lookup"><span data-stu-id="800b3-108">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="800b3-109">运行以下命令以引导主机应用程序（例如 Excel、Word）并在主机应用程序中注册您的外接程序："**npm run sideload**"</span><span class="sxs-lookup"><span data-stu-id="800b3-109">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="800b3-110">另请参阅</span><span class="sxs-lookup"><span data-stu-id="800b3-110">See also</span></span>

- [<span data-ttu-id="800b3-111">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="800b3-111">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="800b3-112">发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="800b3-112">Publish your Office Add-in</span></span>](../publish/publish.md)