---
title: 使 Excel 外接程序与现有 COM 外接程序兼容
description: 启用与与 Excel 外接程序具有相同功能的等效 COM 加载项的兼容性
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 0890e14466a2cd8f5aff2d1bcf307a43cff28127
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628170"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a><span data-ttu-id="d5bb4-103">使您的 Office 外接程序与现有 COM 加载项兼容 (预览)</span><span class="sxs-lookup"><span data-stu-id="d5bb4-103">Make your Office Add-in compatible with an existing COM add-in (preview)</span></span>

<span data-ttu-id="d5bb4-104">如果您有一个现有的 COM 加载项, 则可以在 Excel 加载项中构建等效的功能, 以将解决方案功能扩展到其他平台 (如 online 或 macOS)。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-104">If you have an existing COM add-in, you can build equivalent functionality in your Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="d5bb4-105">但是, Excel 外接程序没有在 COM 加载项中提供的所有功能。你的 COM 加载项可以提供比 Windows 上的 Excel 外接程序更好的体验。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-105">However, Excel add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Excel add-in on Windows.</span></span>

<span data-ttu-id="d5bb4-106">您可以配置 Excel 加载项, 以便在用户的计算机上已安装等效的 COM 加载项时, Office 将运行 COM 加载项, 而不是 Excel 外接程序。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-106">You can configure your Excel add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Excel add-in.</span></span> <span data-ttu-id="d5bb4-107">COM 加载项称为 "等效", 因为 Office 将根据 Windows 上安装的 COM 加载项和 Excel 加载项之间无缝转换。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Excel add-in depending on which is installed on Windows.</span></span>

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="d5bb4-108">在清单中指定等效的 COM 加载项</span><span class="sxs-lookup"><span data-stu-id="d5bb4-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="d5bb4-109">若要启用与现有 COM 加载项的兼容性, 请在 Excel 外接程序的清单中标识等效的 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Excel add-in.</span></span> <span data-ttu-id="d5bb4-110">然后, 在 Windows 上运行时, Office 将使用 COM 加载项, 而不是 Excel 外接程序。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-110">Then Office will use the COM add-in instead of your Excel add-in when running on Windows.</span></span>

<span data-ttu-id="d5bb4-111">`ProgID`指定等效 COM 加载项的。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="d5bb4-112">在安装 COM 加载项时, Office 将使用 COM 加载项 UI, 而不是 Excel 外接程序的 UI。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-112">Office will then use the COM add-in UI instead of your Excel add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="d5bb4-113">下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="d5bb4-114">通常, 出于完整性的考虑, 这两个示例都会在上下文中显示这两个示例。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="d5bb4-115">它们`ProgID` `FileName`分别由各自标识。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="d5bb4-116">有关 XLL 兼容性的详细信息, 请参阅[使您的自定义函数与 xll 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

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

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="d5bb4-117">用户的等效行为</span><span class="sxs-lookup"><span data-stu-id="d5bb4-117">Equivalent behavior for users</span></span>

<span data-ttu-id="d5bb4-118">当在 Excel 外接程序清单中指定了等效的 COM 加载项时, Office 将在安装等效 COM 加载项时禁止在 Windows 上使用 Excel 外接程序的 UI。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-118">When an equivalent COM add-in is specified in the Excel add-in manifest, Office suppresses your Excel add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="d5bb4-119">这不会影响其他平台 (如 online 或 macOS) 上的 Excel 外接程序 UI。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-119">This does not affect your Excel add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="d5bb4-120">Office 仅隐藏功能区按钮, 不会阻止安装。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="d5bb4-121">因此, Excel 外接程序仍将显示在以下 UI 位置:</span><span class="sxs-lookup"><span data-stu-id="d5bb4-121">Therefore your Excel add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="d5bb4-122">在 **"我的外接程序**" 下, 因为它已安装技术。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="d5bb4-123">作为功能区管理器中的条目。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="d5bb4-124">以下方案描述了根据用户获取 Excel 加载项的方式而发生的情况。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-124">The following scenarios describe what happens depending on how the user acquires the Excel add-in.</span></span>

### <a name="appsource-acquisition-of-an-excel-add-in"></a><span data-ttu-id="d5bb4-125">AppSource 获取 Excel 外接程序</span><span class="sxs-lookup"><span data-stu-id="d5bb4-125">AppSource acquisition of an Excel add-in</span></span>

<span data-ttu-id="d5bb4-126">如果用户从 AppSource 下载 Excel 加载项, 并且已安装等效的 COM 加载项, 则 Office 将执行以下操作:</span><span class="sxs-lookup"><span data-stu-id="d5bb4-126">If a user downloads the Excel add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="d5bb4-127">安装 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-127">Install the Excel add-in.</span></span>
2. <span data-ttu-id="d5bb4-128">在功能区中隐藏 Excel 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-128">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="d5bb4-129">为用户显示一个指出 "COM 加载项" 功能区按钮的调用。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-excel-add-in"></a><span data-ttu-id="d5bb4-130">Excel 加载项的集中部署</span><span class="sxs-lookup"><span data-stu-id="d5bb4-130">Centralized deployment of Excel add-in</span></span>

<span data-ttu-id="d5bb4-131">如果管理员使用集中部署将 Excel 加载项部署到其租户, 并且已安装等效的 COM 加载项, 则用户需要先重新启动 Office, 然后他们才会看到任何更改。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-131">If an admin deploys the Excel add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="d5bb4-132">Office 重启后, 将执行以下操作:</span><span class="sxs-lookup"><span data-stu-id="d5bb4-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="d5bb4-133">安装 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-133">Install the Excel add-in.</span></span>
2. <span data-ttu-id="d5bb4-134">在功能区中隐藏 Excel 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-134">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="d5bb4-135">为用户显示一个指出 "COM 加载项" 功能区按钮的调用。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-excel-add-in"></a><span data-ttu-id="d5bb4-136">与嵌入的 Excel 加载项共享的文档</span><span class="sxs-lookup"><span data-stu-id="d5bb4-136">Document shared with embedded Excel add-in</span></span>

<span data-ttu-id="d5bb4-137">如果用户安装了 COM 外接程序, 然后使用嵌入的 Excel 加载项获取共享文档, 然后当他们打开文档时, Office 将执行以下操作:</span><span class="sxs-lookup"><span data-stu-id="d5bb4-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Excel add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="d5bb4-138">提示用户信任 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-138">Prompt the user to trust the Excel add-in.</span></span>
2. <span data-ttu-id="d5bb4-139">如果受信任, 将安装 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-139">If trusted, the Excel add-in will install.</span></span>
3. <span data-ttu-id="d5bb4-140">在功能区中隐藏 Excel 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-140">Hide the Excel add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="d5bb4-141">其他 COM 加载项行为</span><span class="sxs-lookup"><span data-stu-id="d5bb4-141">Other COM add-in behavior</span></span>

<span data-ttu-id="d5bb4-142">如果用户卸载 COM 加载项, 则 Office 将在 Windows 上还原 Excel 外接程序 UI, 以获取等效的已安装 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-142">If a user uninstalls the COM add-in, then Office restores the Excel add-in UI on Windows for the equivalent installed Excel add-in.</span></span>

<span data-ttu-id="d5bb4-143">为 Excel 加载项指定等效的 COM 加载项后, Office 将停止处理 Excel 加载项的更新。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-143">Once you specify an equivalent COM add-in for your Excel add-in, Office stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="d5bb4-144">用户必须卸载 COM 加载项, 才能获取 Excel 外接程序的最新更新。</span><span class="sxs-lookup"><span data-stu-id="d5bb4-144">The user must uninstall the COM add-in order to get the latest updates for the Excel add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="d5bb4-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d5bb4-145">See also</span></span>

- [<span data-ttu-id="d5bb4-146">使自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="d5bb4-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
