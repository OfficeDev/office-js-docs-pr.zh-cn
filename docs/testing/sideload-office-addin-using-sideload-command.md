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
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="d774b-102">使用旁加载命令旁加载 Office 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="d774b-102">Sideload Office Add-ins for testing using the sideload command</span></span>
 
> [!NOTE]
> <span data-ttu-id="d774b-103">本文中所述的旁加载技术仅适用于：</span><span class="sxs-lookup"><span data-stu-id="d774b-103">The sideloading technique described in this article is only valid for:</span></span>
> 
> - <span data-ttu-id="d774b-104">Windows 上运行的 Excel、Word 和 PowerPoint 加载项</span><span class="sxs-lookup"><span data-stu-id="d774b-104">Excel, Word, and PowerPoint add-ins that run on Windows</span></span>
> 
> - <span data-ttu-id="d774b-105">使用[适合于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建并且 package.json 文件中的 `scripts` 部分具有 `sideload` 脚本的加载项项目。</span><span class="sxs-lookup"><span data-stu-id="d774b-105">Add-in projects that were created with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="d774b-106">（使用更早版本的适用于 Office 加载项的 Yeoman 生成器创建的项目没有此脚本。）</span><span class="sxs-lookup"><span data-stu-id="d774b-106">(Projects that were created with older versions of the Yeoman generator for Office Add-ins will not have this script.)</span></span>
 
<span data-ttu-id="d774b-107">若要使用适合于 Office 加载项的 Yeoman 生成器提供的 `sideload` 脚本旁加载你的加载项，请按照以下步骤操作：</span><span class="sxs-lookup"><span data-stu-id="d774b-107">To sideload your add-in by using the `sideload` script that the Yeoman generator for Office Add-ins provides, complete the following steps:</span></span>

1. <span data-ttu-id="d774b-108">以管理员身份打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="d774b-108">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="d774b-109">将目录更改为加载项项目文件夹的根目录。</span><span class="sxs-lookup"><span data-stu-id="d774b-109">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="d774b-110">运行以下命令以在端口 3000 上启动本地 Web 服务器实例，以便为加载项项目提供服务：`npm run start`</span><span class="sxs-lookup"><span data-stu-id="d774b-110">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "`npm run start`"</span></span>

4. <span data-ttu-id="d774b-111">以管理员身份打开第二个命令提示符。</span><span class="sxs-lookup"><span data-stu-id="d774b-111">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="d774b-112">将目录更改为加载项项目文件夹的根目录。</span><span class="sxs-lookup"><span data-stu-id="d774b-112">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="d774b-113">运行以下命令以引导主机应用程序（例如 Excel、Word）并在主机应用程序中注册你的加载项：`npm run sideload`</span><span class="sxs-lookup"><span data-stu-id="d774b-113">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "`npm run sideload`"</span></span>

<span data-ttu-id="d774b-114">如果你的加载项项目是使用 Visual Studio 创建的，或者没有旁加载脚本，则可以使用[从网络共享旁加载 Office 加载项](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中所述的方法在 Windows 上旁加载它。</span><span class="sxs-lookup"><span data-stu-id="d774b-114">(Projects that were created with older versions of yo office also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

<span data-ttu-id="d774b-115">如果不在 Windows 上测试 Word、Excel 或 PowerPoint 加载项，则请参阅以下主题之一，以了解与旁加载你的加载项相关的信息：</span><span class="sxs-lookup"><span data-stu-id="d774b-115">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
 
- [<span data-ttu-id="d774b-116">在 Office Online 中旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="d774b-116">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="d774b-117">在 iPad 和 Mac 上旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="d774b-117">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="d774b-118">旁加载 Outlook 加载项以供测试</span><span class="sxs-lookup"><span data-stu-id="d774b-118">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a><span data-ttu-id="d774b-119">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d774b-119">See also</span></span>

- [<span data-ttu-id="d774b-120">验证并排查清单问题</span><span class="sxs-lookup"><span data-stu-id="d774b-120">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="d774b-121">发布 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="d774b-121">Publish your Office Add-in</span></span>](../publish/publish.md)
