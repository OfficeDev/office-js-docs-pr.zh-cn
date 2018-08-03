---
title: 使用 sideload 命令旁加载 Office 外接程序
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 90084fad0e79ab8acdf59eaa305825737401c0c8
ms.sourcegitcommit: e094aaa06d9aff3d13f8ffd3429d4a31f0b65b81
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/03/2018
ms.locfileid: "21782824"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="0502f-102">旁加载 Office 外接程序，以供使用 **sideload 命令**测试</span><span class="sxs-lookup"><span data-stu-id="0502f-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="0502f-103">“npm run sideload”方法仅适用于在 Windows 上运行的 Excel、Word 和 PowerPoint 加载项；并且仅适用于使用 [**yo office** 工具](https://github.com/OfficeDev/generator-office)创建并在 package.json 文件的 `scripts` 部分有 `sideload` 脚本的加载项项目。</span><span class="sxs-lookup"><span data-stu-id="0502f-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="0502f-104">（使用 **yo office** 旧版本创建的项目也没有这个脚本。）如果你的项目是用 Visual Studio 创建的，或者没有旁加载脚本，你可以用[从网络共享旁加载 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中描述的方法在 Windows 上旁加载它。</span><span class="sxs-lookup"><span data-stu-id="0502f-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="0502f-105">如果不在 Windows 上测试 Word、Excel 或 PowerPoint 加载项，请参阅以下主题之一来旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="0502f-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="0502f-106">在 Office Online 中旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="0502f-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="0502f-107">在 iPad 和 Mac 上旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="0502f-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="0502f-108">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="0502f-108">Sideload Outlook add-ins for testing</span></span>](../../../../outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="0502f-109">以管理员身份打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="0502f-109">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="0502f-110">将目录更改为外接程序项目文件夹的根目录。</span><span class="sxs-lookup"><span data-stu-id="0502f-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="0502f-111">运行以下命令以，在端口 3000 上启动本地 Web 服务器实例以提供外接程序项目："**npm run start**"</span><span class="sxs-lookup"><span data-stu-id="0502f-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="0502f-112">以管理员身份打开第二个命令提示符。</span><span class="sxs-lookup"><span data-stu-id="0502f-112">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="0502f-113">将目录更改为外接程序项目文件夹的根目录。</span><span class="sxs-lookup"><span data-stu-id="0502f-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="0502f-114">运行以下命令以引导主机应用程序（例如 Excel、Word）并在主机应用程序中注册您的外接程序："**npm run sideload**"</span><span class="sxs-lookup"><span data-stu-id="0502f-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="0502f-115">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0502f-115">See also</span></span>

- [<span data-ttu-id="0502f-116">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="0502f-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="0502f-117">发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="0502f-117">Publish your Office Add-in</span></span>](../publish/publish.md)