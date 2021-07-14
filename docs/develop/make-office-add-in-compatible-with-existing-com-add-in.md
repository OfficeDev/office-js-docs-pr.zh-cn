---
title: 确认 Office 加载项与已有的COM 加载项兼容
description: 启用你的Office加载项和等效 COM 加载项之间的兼容性。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 85e5d8cc06aa599862c92b59a26c744f28ca2d22
ms.sourcegitcommit: 95fc1fc8a0dbe8fc94f0ea647836b51cc7f8601d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/14/2021
ms.locfileid: "53418683"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="7fb9a-103">确认 Office 加载项与已有的COM 加载项兼容</span><span class="sxs-lookup"><span data-stu-id="7fb9a-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="7fb9a-104">如果你有现有的 COM 加载项，可以在 Office 加载项中生成等效功能，从而使你的解决方案可以在其他平台（如 Office web 版 或 Mac）中运行。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="7fb9a-105">在某些情况下，Office加载项可能无法提供相应 COM 加载项中提供的所有功能。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="7fb9a-106">在这些情况下，COM 加载项可以提供更好的用户体验，Windows外接程序Office相应的用户体验。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="7fb9a-107">您可以配置 Office 外接程序，以便当用户的计算机上已安装等效 COM 加载项时，Windows 上的 Office 将运行 COM 加载项，而不是 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="7fb9a-108">COM 加载项称为"等效"，因为 Office 将按照安装用户计算机时在 COM 加载项和 Office 加载项之间无缝转换。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="7fb9a-109">当连接到订阅订阅时，以下平台和应用程序Microsoft 365此功能。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-109">This feature is supported by the following platform and applications, when connected to a Microsoft 365 subscription.</span></span> <span data-ttu-id="7fb9a-110">COM 加载项无法安装在任何其他平台上，因此在这些平台上，将忽略本文稍后讨论的清单 `EquivalentAddins` 元素。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-110">COM add-ins cannot be installed on any other platform, so on those platforms the manifest element that is discussed later in this article, `EquivalentAddins`, is ignored.</span></span>
>
> - <span data-ttu-id="7fb9a-111">Excel版本 1904 PowerPoint更高版本Windows (、Word 和) </span><span class="sxs-lookup"><span data-stu-id="7fb9a-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="7fb9a-112">指定等效的 COM 加载项</span><span class="sxs-lookup"><span data-stu-id="7fb9a-112">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="7fb9a-113">清单</span><span class="sxs-lookup"><span data-stu-id="7fb9a-113">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7fb9a-114">适用于 Excel、PowerPoint 和 Word。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-114">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="7fb9a-115">Outlook即将推出支持。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-115">Outlook support coming soon.</span></span>

<span data-ttu-id="7fb9a-116">若要在加载项Office COM 加载项之间实现兼容性，请确定加载项清单中等效的 COM Office加载项。 [](add-in-manifests.md)</span><span class="sxs-lookup"><span data-stu-id="7fb9a-116">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="7fb9a-117">然后Office加载项Windows COM 加载项，而不是Office加载项（如果两者均已安装）。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-117">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="7fb9a-118">以下示例显示清单中将 COM 加载项指定为等效加载项的部分。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-118">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="7fb9a-119">元素的值标识 `ProgId` COM 加载项， [而 EquivalentAddins](../reference/manifest/equivalentaddins.md) 元素必须紧接在结束标记 `VersionOverrides` 的之前。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-119">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

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
> <span data-ttu-id="7fb9a-120">有关 COM 加载项和 XLL UDF 兼容性的信息，请参阅使自定义函数与 [XLL 用户定义函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-120">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="7fb9a-121">组策略</span><span class="sxs-lookup"><span data-stu-id="7fb9a-121">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7fb9a-122">仅适用于Outlook。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-122">Applies to Outlook only.</span></span>

<span data-ttu-id="7fb9a-123">若要声明 Outlook Web 加载项和 COM/VSTO 加载项之间的兼容性，请标识组策略停用 **Outlook Web** 加载项中的等效 COM 加载项，这些加载项的等效 COM 或 VSTO 加载项通过配置安装在用户计算机上。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-123">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="7fb9a-124">然后Outlook加载项Windows COM 加载项，而不是 Web 加载项（如果两者均已安装）。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-124">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="7fb9a-125">下载最新的 [管理模板工具](https://www.microsoft.com/download/details.aspx?id=49030)，注意该工具的 **安装说明**。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-125">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="7fb9a-126">打开 **gpedit.msc (本地组策略**) 。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-126">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="7fb9a-127">导航到 **用户配置**  >  **管理模板**   >  **Microsoft Outlook 2016**  >  **杂项**。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-127">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="7fb9a-128">选择"停用 **Outlook加载项的** 等效 COM 或VSTO Web 加载项"设置。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-128">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="7fb9a-129">打开链接以编辑策略设置。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-129">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="7fb9a-130">在对话框中 **Outlook Web 外接程序停用**：</span><span class="sxs-lookup"><span data-stu-id="7fb9a-130">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="7fb9a-131">将 **"值** `Id` 名称"设置为在 Web 加载项清单中找到的 。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-131">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="7fb9a-132">**重要** 提示 *：请勿* 在条目周围 `{}` 添加大括号。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-132">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="7fb9a-133">将 **"** 值 `ProgId` "设置为等效 COM/VSTO加载项的 。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-133">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="7fb9a-134">选择 **"** 确定"将更新生效。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-134">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="7fb9a-135">![Screenshot showing the dialog "Outlook web add-ins to deactivate".](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="7fb9a-135">![Screenshot showing the dialog "Outlook web add-ins to deactivate".](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="7fb9a-136">用户的等效行为</span><span class="sxs-lookup"><span data-stu-id="7fb9a-136">Equivalent behavior for users</span></span>

<span data-ttu-id="7fb9a-137">如果指定了等效[COM](#specify-an-equivalent-com-add-in)加载项，Windows 上的 Office 将不会显示 Office 加载项的用户界面 (UI) 如果安装了等效的 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-137">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="7fb9a-138">Office仅隐藏加载项的功能Office按钮，不会阻止安装。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-138">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="7fb9a-139">因此Office外接程序仍将显示在 UI 内的以下位置。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-139">Therefore your Office Add-in will still appear in the following locations within the UI.</span></span>

- <span data-ttu-id="7fb9a-140">在 **"我的外接程序"下**</span><span class="sxs-lookup"><span data-stu-id="7fb9a-140">Under **My add-ins**</span></span>
- <span data-ttu-id="7fb9a-141">作为功能区管理器中的条目， (Excel、Word 和 PowerPoint仅) </span><span class="sxs-lookup"><span data-stu-id="7fb9a-141">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="7fb9a-142">在清单中指定等效的 COM 加载项对于其他平台（如 Office web 版 或 Mac）没有影响。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-142">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="7fb9a-143">以下方案描述了根据用户如何获取加载项Office发生的情况。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-143">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="7fb9a-144">AppSource 获取Office加载项</span><span class="sxs-lookup"><span data-stu-id="7fb9a-144">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="7fb9a-145">如果用户从 AppSource Office加载项，并且已安装等效的 COM 加载项，Office将：</span><span class="sxs-lookup"><span data-stu-id="7fb9a-145">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="7fb9a-146">安装Office加载项。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-146">Install the Office Add-in.</span></span>
2. <span data-ttu-id="7fb9a-147">隐藏Office功能区中的加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-147">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="7fb9a-148">为指出 COM 加载项功能区按钮的用户显示一个调用。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-148">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="7fb9a-149">加载项Office集中部署</span><span class="sxs-lookup"><span data-stu-id="7fb9a-149">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="7fb9a-150">如果管理员使用集中式部署将 Office 外接程序部署到其租户，并且已安装等效的 COM 外接程序，则用户必须先重新启动 Office，然后才能看到任何更改。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-150">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="7fb9a-151">重新启动Office，它将：</span><span class="sxs-lookup"><span data-stu-id="7fb9a-151">After Office restarts, it will:</span></span>

1. <span data-ttu-id="7fb9a-152">安装Office加载项。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-152">Install the Office Add-in.</span></span>
2. <span data-ttu-id="7fb9a-153">隐藏Office功能区中的加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-153">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="7fb9a-154">为指出 COM 加载项功能区按钮的用户显示一个调用。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-154">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="7fb9a-155">与嵌入加载项Office的文档</span><span class="sxs-lookup"><span data-stu-id="7fb9a-155">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="7fb9a-156">如果用户已安装 COM 加载项，然后获取与嵌入式 Office 加载项的共享文档，那么当用户打开该文档时，Office将：</span><span class="sxs-lookup"><span data-stu-id="7fb9a-156">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="7fb9a-157">提示用户信任Office外接程序。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-157">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="7fb9a-158">如果受信任，Office外接程序将安装。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-158">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="7fb9a-159">隐藏Office功能区中的加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-159">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="7fb9a-160">其他 COM 加载项行为</span><span class="sxs-lookup"><span data-stu-id="7fb9a-160">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="7fb9a-161">Excel、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="7fb9a-161">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="7fb9a-162">如果用户卸载等效的 COM 加载项，Office加载项WINDOWS会Office加载项 UI。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-162">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="7fb9a-163">为加载项指定等效的 COM Office后，Office停止处理加载项Office更新。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-163">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="7fb9a-164">若要获取加载项的最新Office，用户必须先卸载 COM 加载项。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-164">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="7fb9a-165">Outlook</span><span class="sxs-lookup"><span data-stu-id="7fb9a-165">Outlook</span></span>

<span data-ttu-id="7fb9a-166">COM/VSTO加载项必须在启动Outlook连接，才能禁用相应的 Web 加载项。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-166">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="7fb9a-167">如果 COM/VSTO在后续 Outlook 会话期间断开连接，Web 外接程序可能一直处于禁用状态，直到 Outlook 重新启动。</span><span class="sxs-lookup"><span data-stu-id="7fb9a-167">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="7fb9a-168">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7fb9a-168">See also</span></span>

- [<span data-ttu-id="7fb9a-169">使自定义函数与 XLL 用户定义函数兼容</span><span class="sxs-lookup"><span data-stu-id="7fb9a-169">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
