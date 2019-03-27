---
title: 使用旁加载命令旁加载 Office 加载项
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: dfa231374133ad857554afaf343362f1415788f4
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870112"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="f2574-102">使用**旁加载命令**旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="f2574-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="f2574-103">“npm run sideload”方法仅适用于在 Windows 上运行的 Excel、Word 和 PowerPoint 加载项；并且仅适用于使用 [**yo office** 工具](https://github.com/OfficeDev/generator-office)创建并且在 package.json 文件的 `scripts` 部分中具有 `sideload` 脚本的加载项。</span><span class="sxs-lookup"><span data-stu-id="f2574-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="f2574-104">（使用旧版 **yo office** 创建的项目也没有此脚本。）如果你的项目是使用 Visual Studio 创建的，或者没有旁加载脚本，则可以使用[从网络共享旁加载 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中所述的方法在 Windows 上旁加载它。</span><span class="sxs-lookup"><span data-stu-id="f2574-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="f2574-105">如果不在 Windows 上测试 Word、Excel 或 PowerPoint 外接程序，则请参阅以下主题之一来旁加载外接程序：</span><span class="sxs-lookup"><span data-stu-id="f2574-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="f2574-106">在 Office Online 中旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="f2574-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="f2574-107">在 iPad 和 Mac 上旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="f2574-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="f2574-108">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="f2574-108">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="f2574-109">以管理员身份打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="f2574-109">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="f2574-110">将目录更改为加载项项目文件夹的根目录。</span><span class="sxs-lookup"><span data-stu-id="f2574-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="f2574-111">运行以下命令以在端口 3000 上启动本地 Web 服务器实例，以便为加载项项目提供服务：“**npm run start**”</span><span class="sxs-lookup"><span data-stu-id="f2574-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="f2574-112">以管理员身份打开第二个命令提示符。</span><span class="sxs-lookup"><span data-stu-id="f2574-112">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="f2574-113">将目录更改为加载项项目文件夹的根目录。</span><span class="sxs-lookup"><span data-stu-id="f2574-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="f2574-114">运行以下命令以引导主机应用程序（例如 Excel、Word）并在主机应用程序中注册你的加载项：“**npm run sideload**”</span><span class="sxs-lookup"><span data-stu-id="f2574-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="f2574-115">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f2574-115">See also</span></span>

- [<span data-ttu-id="f2574-116">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="f2574-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="f2574-117">发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="f2574-117">Publish your Office Add-in</span></span>](../publish/publish.md)
