---
title: 确认 Office 加载项与已有的COM 加载项兼容
description: 启用 Office 加载项和等效 COM 加载项之间的兼容性。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b5235255987cc6a644491bc548b92701b350a179
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836850"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="8b29c-103">确认 Office 加载项与已有的COM 加载项兼容</span><span class="sxs-lookup"><span data-stu-id="8b29c-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="8b29c-104">如果你有现有的 COM 加载项，可以在 Office 加载项中生成等效功能，从而使你的解决方案可以在其他平台（如 Office 网页或 Mac）中运行。</span><span class="sxs-lookup"><span data-stu-id="8b29c-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="8b29c-105">在某些情况下，Office 加载项可能无法提供相应 COM 加载项中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="8b29c-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="8b29c-106">在这些情况下，COM 加载项在 Windows 上的用户体验可能会优于相应的 Office 加载项提供的用户体验。</span><span class="sxs-lookup"><span data-stu-id="8b29c-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="8b29c-107">您可以配置 Office 外接程序，以便当用户的计算机上已安装等效 COM 加载项时，Windows 上的 Office 将运行 COM 加载项，而不是 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8b29c-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="8b29c-108">COM 加载项称为"等效"加载项，因为 Office 将在 COM 加载项和 Office 加载项之间无缝转换，根据加载项安装用户计算机。</span><span class="sxs-lookup"><span data-stu-id="8b29c-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="8b29c-109">连接到 Microsoft 365 订阅时，以下平台支持此功能。</span><span class="sxs-lookup"><span data-stu-id="8b29c-109">This feature is supported by the following platforms, when connected to a Microsoft 365 subscription.</span></span>
>
> - <span data-ttu-id="8b29c-110">Excel、Word 和 PowerPoint 网页</span><span class="sxs-lookup"><span data-stu-id="8b29c-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="8b29c-111">Windows 版 Excel、Word 和 PowerPoint (版本 1904 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="8b29c-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="8b29c-112">Mac 版 Excel、Word 和 PowerPoint (版本 13.329 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="8b29c-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>
> - <span data-ttu-id="8b29c-113">Windows 版 Outlook (版本 2102 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="8b29c-113">Outlook on Windows (version 2102 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="8b29c-114">指定等效的 COM 加载项</span><span class="sxs-lookup"><span data-stu-id="8b29c-114">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="8b29c-115">清单</span><span class="sxs-lookup"><span data-stu-id="8b29c-115">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8b29c-116">适用于 Excel、PowerPoint 和 Word。</span><span class="sxs-lookup"><span data-stu-id="8b29c-116">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="8b29c-117">即将推出 Outlook 支持。</span><span class="sxs-lookup"><span data-stu-id="8b29c-117">Outlook support coming soon.</span></span>

<span data-ttu-id="8b29c-118">若要启用 Office 加载项和 COM 加载项之间的兼容性，请确定 Office 加载项清单中的等效[](add-in-manifests.md)COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="8b29c-118">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="8b29c-119">然后，Windows 上的 Office 将使用 COM 加载项，而不是 Office 加载项（如果两者均已安装）。</span><span class="sxs-lookup"><span data-stu-id="8b29c-119">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="8b29c-120">以下示例显示清单中将 COM 加载项指定为等效加载项的部分。</span><span class="sxs-lookup"><span data-stu-id="8b29c-120">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="8b29c-121">元素的值标识 `ProgId` COM 加载项， [而 EquivalentAddins](../reference/manifest/equivalentaddins.md) 元素必须紧接在结束标记 `VersionOverrides` 的之前。</span><span class="sxs-lookup"><span data-stu-id="8b29c-121">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="8b29c-122">有关 COM 加载项和 XLL UDF 兼容性的信息，请参阅使自定义函数与 [XLL 用户定义函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="8b29c-122">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="8b29c-123">组策略</span><span class="sxs-lookup"><span data-stu-id="8b29c-123">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8b29c-124">仅适用于 Outlook。</span><span class="sxs-lookup"><span data-stu-id="8b29c-124">Applies to Outlook only.</span></span>

<span data-ttu-id="8b29c-125">若要声明 Outlook Web 加载项和 COM/VSTO 加载项之间的兼容性，请标识组策略停用其等效 COM 或 VSTO 加载项的等效 COM 或 **VSTO** 加载项的组策略中的等效 COM 加载项，在用户计算机上进行配置。</span><span class="sxs-lookup"><span data-stu-id="8b29c-125">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="8b29c-126">然后，Windows 上的 Outlook 将使用 COM 加载项，而不是 Web 加载项（如果两者均已安装）。</span><span class="sxs-lookup"><span data-stu-id="8b29c-126">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="8b29c-127">下载最新的 [管理模板工具](https://www.microsoft.com/download/details.aspx?id=49030)，注意该工具的 **安装说明**。</span><span class="sxs-lookup"><span data-stu-id="8b29c-127">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="8b29c-128">打开 **gpedit.msc (本地组策略**) 。</span><span class="sxs-lookup"><span data-stu-id="8b29c-128">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="8b29c-129">导航到 **"用户配置**  >  **""管理模板**   >  **""Microsoft Outlook 2016**  >  **杂项"。**</span><span class="sxs-lookup"><span data-stu-id="8b29c-129">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="8b29c-130">选择"停用安装了等效 COM 或 VSTO 加载项的 **Outlook Web 加载项"设置**。</span><span class="sxs-lookup"><span data-stu-id="8b29c-130">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="8b29c-131">打开链接以编辑策略设置。</span><span class="sxs-lookup"><span data-stu-id="8b29c-131">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="8b29c-132">在对话框中 **，Outlook Web 外接程序将停用**：</span><span class="sxs-lookup"><span data-stu-id="8b29c-132">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="8b29c-133">将 **"值** `Id` 名称"设置为在 Web 加载项清单中找到的 。</span><span class="sxs-lookup"><span data-stu-id="8b29c-133">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="8b29c-134">**重要** 提示 *：请勿* 在条目周围 `{}` 添加大括号。</span><span class="sxs-lookup"><span data-stu-id="8b29c-134">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="8b29c-135">将 **"** 值 `ProgId` "设置为等效 COM/VSTO 加载项的 。</span><span class="sxs-lookup"><span data-stu-id="8b29c-135">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="8b29c-136">选择 **"** 确定"将更新生效。</span><span class="sxs-lookup"><span data-stu-id="8b29c-136">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="8b29c-137">![Screenshot showing the dialog "Outlook web add-ins to deactivate"](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="8b29c-137">![Screenshot showing the dialog "Outlook web add-ins to deactivate"](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="8b29c-138">用户的等效行为</span><span class="sxs-lookup"><span data-stu-id="8b29c-138">Equivalent behavior for users</span></span>

<span data-ttu-id="8b29c-139">指定等效 COM 加载项时，如果安装了等效 [的 COM](#specify-an-equivalent-com-add-in)加载项，Windows 上的 Office (UI) Office 加载项的用户界面将不会显示。</span><span class="sxs-lookup"><span data-stu-id="8b29c-139">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="8b29c-140">Office 仅隐藏 Office 外接程序的功能区按钮，不会阻止安装。</span><span class="sxs-lookup"><span data-stu-id="8b29c-140">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="8b29c-141">因此，您的 Office 外接程序仍将显示在 UI 内的以下位置：</span><span class="sxs-lookup"><span data-stu-id="8b29c-141">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="8b29c-142">在 **"我的外接程序"下**</span><span class="sxs-lookup"><span data-stu-id="8b29c-142">Under **My add-ins**</span></span>
- <span data-ttu-id="8b29c-143">作为功能区管理器中的条目， (Excel、Word 和 PowerPoint) </span><span class="sxs-lookup"><span data-stu-id="8b29c-143">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="8b29c-144">在清单中指定等效 COM 加载项对 Web 上的 Office 或 Mac 上的其他平台没有影响。</span><span class="sxs-lookup"><span data-stu-id="8b29c-144">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="8b29c-145">以下方案描述了根据用户如何获取 Office 外接程序而发生的情况。</span><span class="sxs-lookup"><span data-stu-id="8b29c-145">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="8b29c-146">Office 加载项的 AppSource 购置</span><span class="sxs-lookup"><span data-stu-id="8b29c-146">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="8b29c-147">如果用户从 AppSource 获取 Office 加载项，并且已安装等效 COM 加载项，则 Office 将：</span><span class="sxs-lookup"><span data-stu-id="8b29c-147">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="8b29c-148">安装 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8b29c-148">Install the Office Add-in.</span></span>
2. <span data-ttu-id="8b29c-149">在功能区中隐藏 Office 外接程序 UI。</span><span class="sxs-lookup"><span data-stu-id="8b29c-149">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="8b29c-150">为指出 COM 加载项功能区按钮的用户显示一个调用。</span><span class="sxs-lookup"><span data-stu-id="8b29c-150">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="8b29c-151">Office 加载项的集中部署</span><span class="sxs-lookup"><span data-stu-id="8b29c-151">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="8b29c-152">如果管理员使用集中式部署将 Office 外接程序部署到其租户，并且已安装等效的 COM 加载项，则用户必须先重新启动 Office，然后才能看到任何更改。</span><span class="sxs-lookup"><span data-stu-id="8b29c-152">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="8b29c-153">Office 重新启动后，它将：</span><span class="sxs-lookup"><span data-stu-id="8b29c-153">After Office restarts, it will:</span></span>

1. <span data-ttu-id="8b29c-154">安装 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="8b29c-154">Install the Office Add-in.</span></span>
2. <span data-ttu-id="8b29c-155">在功能区中隐藏 Office 外接程序 UI。</span><span class="sxs-lookup"><span data-stu-id="8b29c-155">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="8b29c-156">为指出 COM 加载项功能区按钮的用户显示一个调用。</span><span class="sxs-lookup"><span data-stu-id="8b29c-156">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="8b29c-157">与嵌入式 Office 外接程序共享的文档</span><span class="sxs-lookup"><span data-stu-id="8b29c-157">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="8b29c-158">如果用户已安装 COM 加载项，然后获取与嵌入的 Office 加载项的共享文档，则当打开该文档时，Office 将：</span><span class="sxs-lookup"><span data-stu-id="8b29c-158">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="8b29c-159">提示用户信任 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="8b29c-159">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="8b29c-160">如果受信任，将安装 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="8b29c-160">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="8b29c-161">在功能区中隐藏 Office 外接程序 UI。</span><span class="sxs-lookup"><span data-stu-id="8b29c-161">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="8b29c-162">其他 COM 加载项行为</span><span class="sxs-lookup"><span data-stu-id="8b29c-162">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="8b29c-163">Excel、PowerPoint 和 Word</span><span class="sxs-lookup"><span data-stu-id="8b29c-163">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="8b29c-164">如果用户卸载等效的 COM 加载项，Windows 上的 Office 将还原 Office 加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="8b29c-164">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="8b29c-165">为 Office 加载项指定等效 COM 加载项后，Office 将停止处理 Office 加载项的更新。</span><span class="sxs-lookup"><span data-stu-id="8b29c-165">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="8b29c-166">若要获取 Office 加载项的最新更新，用户必须先卸载 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="8b29c-166">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="8b29c-167">Outlook</span><span class="sxs-lookup"><span data-stu-id="8b29c-167">Outlook</span></span>

<span data-ttu-id="8b29c-168">启动 Outlook 时，COM/VSTO 加载项必须连接，才能禁用相应的 Web 加载项。</span><span class="sxs-lookup"><span data-stu-id="8b29c-168">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="8b29c-169">如果在后续 Outlook 会话期间 COM/VSTO 加载项断开连接，在重新启动 Outlook 之前，Web 加载项可能保持禁用状态。</span><span class="sxs-lookup"><span data-stu-id="8b29c-169">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="8b29c-170">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8b29c-170">See also</span></span>

- [<span data-ttu-id="8b29c-171">使自定义函数与 XLL 用户定义函数兼容</span><span class="sxs-lookup"><span data-stu-id="8b29c-171">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
