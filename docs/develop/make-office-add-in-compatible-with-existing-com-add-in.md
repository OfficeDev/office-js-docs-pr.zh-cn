---
title: 让 Office 加载项与现有 COM 加载项兼容
description: 启用 Office 加载项和等效 COM 加载项之间的兼容性
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: a18adb9841a9580d77c5110a0346f365e38e3746
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059718"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a><span data-ttu-id="4157f-103">使您的 Office 外接程序与现有 COM 加载项兼容 (预览)</span><span class="sxs-lookup"><span data-stu-id="4157f-103">Make your Office Add-in compatible with an existing COM add-in (preview)</span></span>

<span data-ttu-id="4157f-104">如果您有一个现有的 COM 加载项, 则可以在 Office 加载项中构建等效功能, 从而使您的解决方案能够在其他平台 (如 web 或 Mac 上的 Office) 上运行。</span><span class="sxs-lookup"><span data-stu-id="4157f-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Office on Mac.</span></span> <span data-ttu-id="4157f-105">在某些情况下, Office 外接程序可能无法提供相应 COM 外接程序中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="4157f-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="4157f-106">在这些情况下, 您的 COM 外接程序在 Windows 上提供的用户体验可能比相应的 Office 外接程序提供的更好。</span><span class="sxs-lookup"><span data-stu-id="4157f-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="4157f-107">您可以配置 Office 加载项, 以便在用户的计算机上已安装等效的 COM 加载项时, Windows 上的 Office 将运行 COM 加载项, 而不是 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="4157f-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="4157f-108">COM 加载项称为 "等效", 因为 Office 将根据安装了用户计算机的加载项和 Office 加载项在 COM 加载项之间进行无缝转换。</span><span class="sxs-lookup"><span data-stu-id="4157f-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="4157f-109">此功能当前处于预览阶段, 不受支持在生产环境中使用。</span><span class="sxs-lookup"><span data-stu-id="4157f-109">This feature is currently in preview and not supported for use in production environments.</span></span> <span data-ttu-id="4157f-110">它在 Excel、Word 和 PowerPoint 版本16.0.11629.20214 或更高版本中可用。</span><span class="sxs-lookup"><span data-stu-id="4157f-110">It's available in Excel, Word, and PowerPoint version 16.0.11629.20214 or later.</span></span> <span data-ttu-id="4157f-111">若要访问此版本, 您必须拥有 Office 365 订阅, 并在**内幕**级加入[Office 预览体验成员](https://products.office.com/office-insider)计划。</span><span class="sxs-lookup"><span data-stu-id="4157f-111">To access this build, you must have an Office 365 subscription and join the [Office Insider](https://products.office.com/office-insider) program at the **Insider** level.</span></span>

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="4157f-112">在清单中指定等效的 COM 加载项</span><span class="sxs-lookup"><span data-stu-id="4157f-112">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="4157f-113">若要在 Office 外接程序和 COM 加载项之间启用兼容性, 请在 Office 外接程序的[清单](add-in-manifests.md)中标识等效的 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="4157f-113">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="4157f-114">然后, Windows 上的 Office 将使用 COM 加载项, 而不是 Office 加载项 (如果已安装)。</span><span class="sxs-lookup"><span data-stu-id="4157f-114">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="4157f-115">以下示例显示了将 COM 加载项指定为等效加载项的清单部分。</span><span class="sxs-lookup"><span data-stu-id="4157f-115">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="4157f-116">`ProgId`元素的值标识 COM 加载项, 并且`EquivalentAddins`元素必须紧跟在结束`VersionOverrides`标记之前。</span><span class="sxs-lookup"><span data-stu-id="4157f-116">The value of the `ProgId` element identifies the COM add-in and the `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  <EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="4157f-117">有关 COM 加载项和 XLL UDF 兼容性的信息, 请参阅[使您的自定义函数与 XLL 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="4157f-117">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="4157f-118">用户的等效行为</span><span class="sxs-lookup"><span data-stu-id="4157f-118">Equivalent behavior for users</span></span>

<span data-ttu-id="4157f-119">在 Office 外接程序清单中指定等效的 COM 外接程序时, 如果安装了等效的 COM 加载项, 则 Windows 上的 Office 将不会显示 Office 加载项的用户界面 (UI)。</span><span class="sxs-lookup"><span data-stu-id="4157f-119">When an equivalent COM add-in is specified in the Office Add-in manifest, Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="4157f-120">Office 仅隐藏 Office 加载项的功能区按钮, 不会阻止安装。</span><span class="sxs-lookup"><span data-stu-id="4157f-120">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="4157f-121">因此, 你的 Office 外接程序仍将显示在 UI 中的以下位置:</span><span class="sxs-lookup"><span data-stu-id="4157f-121">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="4157f-122">在 **"我的外接程序**" 下</span><span class="sxs-lookup"><span data-stu-id="4157f-122">Under **My add-ins**</span></span>
- <span data-ttu-id="4157f-123">作为功能区管理器中的条目</span><span class="sxs-lookup"><span data-stu-id="4157f-123">As an entry in the ribbon manager</span></span>

> [!NOTE]
> <span data-ttu-id="4157f-124">在清单中指定等效的 COM 加载项不会对其他平台 (如 web 上的 Office 或 Office for Mac) 产生影响。</span><span class="sxs-lookup"><span data-stu-id="4157f-124">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or Office for Mac.</span></span>

<span data-ttu-id="4157f-125">以下方案描述了根据用户获取 Office 加载项的方式而发生的情况。</span><span class="sxs-lookup"><span data-stu-id="4157f-125">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="4157f-126">AppSource Office 外接程序的获取</span><span class="sxs-lookup"><span data-stu-id="4157f-126">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="4157f-127">如果用户从 AppSource 获取 Office 加载项, 并且已安装等效的 COM 加载项, 则 Office 将:</span><span class="sxs-lookup"><span data-stu-id="4157f-127">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="4157f-128">安装 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4157f-128">Install the Office Add-in.</span></span>
2. <span data-ttu-id="4157f-129">在功能区中隐藏 Office 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="4157f-129">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="4157f-130">为用户显示一个指出 "COM 加载项" 功能区按钮的调用。</span><span class="sxs-lookup"><span data-stu-id="4157f-130">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="4157f-131">Office 加载项的集中部署</span><span class="sxs-lookup"><span data-stu-id="4157f-131">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="4157f-132">如果管理员使用集中部署将 Office 加载项部署到其租户, 并且已安装了等效的 COM 加载项, 则用户必须重新启动 Office 才能看到任何更改。</span><span class="sxs-lookup"><span data-stu-id="4157f-132">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="4157f-133">Office 重启后, 将执行以下操作:</span><span class="sxs-lookup"><span data-stu-id="4157f-133">After Office restarts, it will:</span></span>

1. <span data-ttu-id="4157f-134">安装 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4157f-134">Install the Office Add-in.</span></span>
2. <span data-ttu-id="4157f-135">在功能区中隐藏 Office 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="4157f-135">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="4157f-136">为用户显示一个指出 "COM 加载项" 功能区按钮的调用。</span><span class="sxs-lookup"><span data-stu-id="4157f-136">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="4157f-137">与嵌入的 Office 加载项共享的文档</span><span class="sxs-lookup"><span data-stu-id="4157f-137">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="4157f-138">如果用户安装了 COM 加载项, 然后使用嵌入的 Office 外接程序获取共享文档, 然后当他们打开文档时, Office 将执行以下操作:</span><span class="sxs-lookup"><span data-stu-id="4157f-138">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="4157f-139">提示用户信任 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="4157f-139">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="4157f-140">如果受信任, Office 加载项将会安装。</span><span class="sxs-lookup"><span data-stu-id="4157f-140">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="4157f-141">在功能区中隐藏 Office 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="4157f-141">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="4157f-142">其他 COM 加载项行为</span><span class="sxs-lookup"><span data-stu-id="4157f-142">Other COM add-in behavior</span></span>

<span data-ttu-id="4157f-143">如果用户卸载等效的 COM 加载项, 则 Windows 上的 Office 将还原 Office 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="4157f-143">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="4157f-144">为 Office 外接程序指定等效的 COM 外接程序后, Office 将停止处理 Office 外接程序的更新。</span><span class="sxs-lookup"><span data-stu-id="4157f-144">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="4157f-145">若要获取 Office 外接程序的最新更新, 用户必须先卸载 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="4157f-145">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="4157f-146">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4157f-146">See also</span></span>

- [<span data-ttu-id="4157f-147">使自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="4157f-147">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
