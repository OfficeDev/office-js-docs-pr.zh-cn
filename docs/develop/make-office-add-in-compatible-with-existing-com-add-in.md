---
title: 使您的 Office 外接程序与现有的 COM 外接程序兼容
description: 启用与与 Office 外接程序具有相同功能的等效 COM 加载项的兼容性
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356847"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="f987e-103">使您的 Office 外接程序与现有的 COM 外接程序兼容</span><span class="sxs-lookup"><span data-stu-id="f987e-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="f987e-104">如果有现有的 COM 加载项, 则可以在 Office 外接程序中生成等效功能, 以将解决方案功能扩展到其他平台 (如 online 或 macOS)。</span><span class="sxs-lookup"><span data-stu-id="f987e-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="f987e-105">但是, Office 外接程序没有在 COM 加载项中提供的所有功能。在 Excel、Word 和 PowerPoint 中, 您的 COM 加载项可以提供比 Windows 上的 Office 外接程序更好的体验。</span><span class="sxs-lookup"><span data-stu-id="f987e-105">However, Office Add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Office Add-in on Windows in Excel, Word, and PowerPoint.</span></span>

<span data-ttu-id="f987e-106">您可以配置 office 加载项, 以便在用户的计算机上已安装等效的 COM 加载项时, office 将运行 COM 加载项, 而不是 office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="f987e-106">You can configure your Office Add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Office Add-in.</span></span> <span data-ttu-id="f987e-107">com 加载项称为 "等效", 因为 Office 将在 COM 加载项和 Office 加载项之间进行无缝转换, 具体取决于 Windows 上安装的版本。</span><span class="sxs-lookup"><span data-stu-id="f987e-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in depending on which is installed on Windows.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="f987e-108">在清单中指定等效的 COM 加载项</span><span class="sxs-lookup"><span data-stu-id="f987e-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="f987e-109">若要启用与现有 COM 加载项的兼容性, 请在 Office 外接程序的清单中标识等效的 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="f987e-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Office Add-in.</span></span> <span data-ttu-id="f987e-110">在 Windows 上运行时, office 将使用 COM 加载项, 而不是 office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="f987e-110">Then Office will use the COM add-in instead of your Office Add-in when running on Windows.</span></span>

<span data-ttu-id="f987e-111">`ProgID`指定等效 COM 加载项的。</span><span class="sxs-lookup"><span data-stu-id="f987e-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="f987e-112">然后, 在安装 com 加载项时, office 将使用 com 加载项 ui, 而不是 office 外接程序的 ui。</span><span class="sxs-lookup"><span data-stu-id="f987e-112">Office will then use the COM add-in UI instead of your Office Add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="f987e-113">下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。</span><span class="sxs-lookup"><span data-stu-id="f987e-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="f987e-114">通常, 出于完整性的考虑, 这两个示例都会在上下文中显示这两个示例。</span><span class="sxs-lookup"><span data-stu-id="f987e-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="f987e-115">它们`ProgID` `FileName`分别由各自标识。</span><span class="sxs-lookup"><span data-stu-id="f987e-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="f987e-116">有关 xll 兼容性的详细信息, 请参阅[使您的自定义函数与 xll 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="f987e-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="f987e-117">用户的等效行为</span><span class="sxs-lookup"><span data-stu-id="f987e-117">Equivalent behavior for users</span></span>

<span data-ttu-id="f987e-118">当 office 外接程序清单中指定了等效的 com 加载项时, office 将在安装等效 com 加载项时在 Windows 上取消使用 office 外接程序的 UI。</span><span class="sxs-lookup"><span data-stu-id="f987e-118">When an equivalent COM add-in is specified in the Office Add-in manifest, Office suppresses your Office Add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="f987e-119">这不会影响其他平台 (如 online 或 macOS) 上的 Office 外接程序的 UI。</span><span class="sxs-lookup"><span data-stu-id="f987e-119">This does not affect your Office Add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="f987e-120">Office 仅隐藏功能区按钮, 不会阻止安装。</span><span class="sxs-lookup"><span data-stu-id="f987e-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="f987e-121">因此, 你的 Office 外接程序仍将显示在以下 UI 位置:</span><span class="sxs-lookup"><span data-stu-id="f987e-121">Therefore your Office Add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="f987e-122">在 **"我的外接程序**" 下, 因为它已安装技术。</span><span class="sxs-lookup"><span data-stu-id="f987e-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="f987e-123">作为功能区管理器中的条目。</span><span class="sxs-lookup"><span data-stu-id="f987e-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="f987e-124">以下方案描述了根据用户获取 Office 加载项的方式而发生的情况。</span><span class="sxs-lookup"><span data-stu-id="f987e-124">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="f987e-125">AppSource Office 外接程序的获取</span><span class="sxs-lookup"><span data-stu-id="f987e-125">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="f987e-126">如果用户从 AppSource 下载 Office 加载项, 并且已安装了等效的 COM 加载项, 则 Office 将执行以下操作:</span><span class="sxs-lookup"><span data-stu-id="f987e-126">If a user downloads the Office Add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="f987e-127">安装 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="f987e-127">Install the Office Add-in.</span></span>
2. <span data-ttu-id="f987e-128">在功能区中隐藏 Office 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="f987e-128">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="f987e-129">为用户显示一个指出 "COM 加载项" 功能区按钮的调用。</span><span class="sxs-lookup"><span data-stu-id="f987e-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="f987e-130">Office 加载项的集中部署</span><span class="sxs-lookup"><span data-stu-id="f987e-130">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="f987e-131">如果管理员使用集中部署将 office 外接程序部署到其租户, 并且已安装等效的 COM 加载项, 则用户需要先重新启动 office, 然后他们才会看到任何更改。</span><span class="sxs-lookup"><span data-stu-id="f987e-131">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="f987e-132">Office 重启后, 将执行以下操作:</span><span class="sxs-lookup"><span data-stu-id="f987e-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="f987e-133">安装 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="f987e-133">Install the Office Add-in.</span></span>
2. <span data-ttu-id="f987e-134">在功能区中隐藏 Office 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="f987e-134">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="f987e-135">为用户显示一个指出 "COM 加载项" 功能区按钮的调用。</span><span class="sxs-lookup"><span data-stu-id="f987e-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="f987e-136">与嵌入的 Office 加载项共享的文档</span><span class="sxs-lookup"><span data-stu-id="f987e-136">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="f987e-137">如果用户安装了 COM 加载项, 然后使用嵌入的 Office 外接程序获取共享文档, 然后当他们打开文档时, Office 将执行以下操作:</span><span class="sxs-lookup"><span data-stu-id="f987e-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="f987e-138">提示用户信任 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="f987e-138">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="f987e-139">如果受信任, Office 加载项将会安装。</span><span class="sxs-lookup"><span data-stu-id="f987e-139">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="f987e-140">在功能区中隐藏 Office 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="f987e-140">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="f987e-141">其他 COM 加载项行为</span><span class="sxs-lookup"><span data-stu-id="f987e-141">Other COM add-in behavior</span></span>

<span data-ttu-id="f987e-142">如果用户卸载 COM 加载项, 则 office 将在 Windows 上还原 office 外接程序 UI, 以获取等效的已安装 office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="f987e-142">If a user uninstalls the COM add-in, then Office restores the Office Add-in UI on Windows for the equivalent installed Office Add-in.</span></span>

<span data-ttu-id="f987e-143">为 office 外接程序指定等效 COM 外接程序后, office 将停止处理 office 外接程序的更新。</span><span class="sxs-lookup"><span data-stu-id="f987e-143">Once you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="f987e-144">用户必须卸载 COM 加载项, 才能获取 Office 外接程序的最新更新。</span><span class="sxs-lookup"><span data-stu-id="f987e-144">The user must uninstall the COM add-in order to get the latest updates for the Office Add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="f987e-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f987e-145">See also</span></span>

- [<span data-ttu-id="f987e-146">使自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="f987e-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
