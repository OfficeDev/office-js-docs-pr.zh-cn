---
title: 将内置 Office 按钮集成到自定义控件组和选项卡中
description: 了解如何在 Office 功能区上的自定义命令组和选项卡中包含内置 Office 按钮。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: e04107893b3c0dd453c84d38fdd5623e308b70e3
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088164"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a><span data-ttu-id="fbcd0-103">将内置 Office 按钮集成到自定义控件组和选项卡中 (预览) </span><span class="sxs-lookup"><span data-stu-id="fbcd0-103">Integrate built-in Office buttons into custom control groups and tabs (preview)</span></span>

<span data-ttu-id="fbcd0-104">您可以使用外接程序清单中的标记将内置 Office 按钮插入 Office 功能区上的自定义控件组中。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="fbcd0-105"> (无法将自定义外接程序命令插入内置 Office 组。 ) 也可以将整个内置 Office 控件组插入到自定义功能区选项卡中。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="fbcd0-106">本文假定您熟悉文章 [外接程序命令的基本概念](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="fbcd0-107">如果你最近未执行此操作，请对其进行检查。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="fbcd0-108">本文中介绍的加载项功能和标记位于预览中， *仅适用于 web 上的 PowerPoint*。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-108">The add-in feature and markup described in this article is in preview and is *only available in PowerPoint on the web*.</span></span> <span data-ttu-id="fbcd0-109">我们建议您仅在测试和开发环境中尝试标记。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-109">We recommend that you try out the markup in test and development environments only.</span></span> <span data-ttu-id="fbcd0-110">请勿在生产环境中或在业务关键型文档中使用预览标记。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-110">Do not use preview markup in a production environment or within business-critical documents.</span></span>
> - <span data-ttu-id="fbcd0-111">本文中所述的标记仅适用于支持要求集 **addincommand 1.3** 的平台。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-111">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="fbcd0-112">有关 [不受支持的平台](#behavior-on-unsupported-platforms)，请参阅后续章节行为。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-112">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="fbcd0-113">在自定义选项卡中插入内置控件组</span><span class="sxs-lookup"><span data-stu-id="fbcd0-113">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="fbcd0-114">若要在选项卡中插入内置 Office 控件组，请在 parent 元素中将 [OfficeGroup](../reference/manifest/customtab.md#officegroup) 元素添加为子元素 `<CustomTab>` 。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-114">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="fbcd0-115">将 `id` 元素的属性 `<OfficeGroup>` 设置为内置组的 ID。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-115">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="fbcd0-116">请参阅 [查找控件和控件组的 id](#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-116">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="fbcd0-117">下面的标记示例将 "Office 段落" 控件组添加到自定义选项卡，并将其放置在自定义组的紧后面。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-117">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="fbcd0-118">将内置控件插入到自定义组中</span><span class="sxs-lookup"><span data-stu-id="fbcd0-118">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="fbcd0-119">若要将内置 Office 控件插入到自定义组中，请在父元素中将 [OfficeControl](../reference/manifest/group.md#officecontrol) 元素添加为子元素 `<Group>` 。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-119">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="fbcd0-120">`id`元素的属性 `<OfficeControl>` 设置为内置控件的 ID。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-120">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="fbcd0-121">请参阅 [查找控件和控件组的 id](#find-the-ids-of-controls-and-control-groups)。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-121">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="fbcd0-122">下面的标记示例将 Office 上标控件添加到自定义组，并将其放置在自定义按钮的紧后面。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-122">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.grp1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

> [!NOTE]
> <span data-ttu-id="fbcd0-123">用户可以在 Office 应用程序中自定义功能区。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-123">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="fbcd0-124">任何用户自定义设置都将覆盖您的清单设置。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-124">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="fbcd0-125">例如，用户可以从任何组中删除按钮，并从选项卡中删除任何组。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-125">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="fbcd0-126">查找控件和控件组的 Id</span><span class="sxs-lookup"><span data-stu-id="fbcd0-126">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="fbcd0-127">支持的控件和控件组的 Id 位于存储库 [Office 控件 id](https://github.com/OfficeDev/office-control-ids)的文件中。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-127">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="fbcd0-128">按照该存储库的自述文件中的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-128">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="fbcd0-129">不受支持的平台上的行为</span><span class="sxs-lookup"><span data-stu-id="fbcd0-129">Behavior on unsupported platforms</span></span>

<span data-ttu-id="fbcd0-130">如果外接程序安装在不支持 [要求集 addincommand 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则会忽略本文中所述的标记，并且内置的 Office 控件/组将不会显示在自定义组/选项卡中。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-130">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="fbcd0-131">若要防止外接程序安装在不支持标记的平台上，请在清单的部分中添加对要求集的引用 `<Requirements>` 。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-131">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="fbcd0-132">有关说明，请参阅 [在清单中设置需求元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-132">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="fbcd0-133">此外，还可以设计外接程序，使其在 **addincommand 1.3** 不受支持时具有备用体验，如 [JavaScript 代码中的使用运行时检查](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)中所述。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-133">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="fbcd0-134">例如，如果您的外接程序包含假设内置按钮位于您的自定义组中的说明，则可以使用一个替代版本，它假定内置按钮仅位于其通常的位置。</span><span class="sxs-lookup"><span data-stu-id="fbcd0-134">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
