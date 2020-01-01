---
title: 清除 Office 缓存
description: 了解如何清除计算机上的 Office 缓存。
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 3744d8125a5165569c262dc28622614853798c6f
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915048"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="6e911-103">清除 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="6e911-103">Clear the Office cache</span></span>

<span data-ttu-id="6e911-104">你可以通过清除计算机上的 Office 缓存来删除以前在 Windows、Mac 或 iOS 上旁加载的加载项。</span><span class="sxs-lookup"><span data-stu-id="6e911-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span> 

<span data-ttu-id="6e911-105">此外，如果你对加载项的清单进行了更改（例如，更新图标的文件名或加载项命令的文本），则应清除 Office 缓存，然后使用更新后的清单重新旁加载此加载项。</span><span class="sxs-lookup"><span data-stu-id="6e911-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="6e911-106">执行此操作后，Office 将按照更新清单中所述的方式呈现该加载项。</span><span class="sxs-lookup"><span data-stu-id="6e911-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="6e911-107">清除 Windows 上的 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="6e911-107">Clear the Office cache on Windows</span></span>

<span data-ttu-id="6e911-108">若要清除 Windows 上的 Office 缓存，请删除文件夹 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` 的内容。</span><span class="sxs-lookup"><span data-stu-id="6e911-108">To clear the Office cache on Windows, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="6e911-109">清除 Mac 上的 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="6e911-109">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="6e911-110">清除 iOS 上的 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="6e911-110">Clear the Office cache on iOS</span></span>

<span data-ttu-id="6e911-111">若要清除 iOS 上的 Office 缓存，请从加载项中的 JavaScript 调用 `window.location.reload(true)` 以强制重新加载。</span><span class="sxs-lookup"><span data-stu-id="6e911-111">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="6e911-112">或者，可以重新安装 Office。</span><span class="sxs-lookup"><span data-stu-id="6e911-112">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="6e911-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6e911-113">See also</span></span>

- [<span data-ttu-id="6e911-114">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="6e911-114">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="6e911-115">验证 Office 加载项的清单</span><span class="sxs-lookup"><span data-stu-id="6e911-115">Validate an Office Add-in manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="6e911-116">使用运行时日志记录功能调试加载项</span><span class="sxs-lookup"><span data-stu-id="6e911-116">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="6e911-117">旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="6e911-117">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="6e911-118">调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="6e911-118">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)