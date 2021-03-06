---
title: 将内置 Office 按钮集成到自定义控件组和选项卡中
description: 了解如何在自定义命令组和 Office 功能区上的选项卡中包括内置 Office 按钮。
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 8d4e8f39313551d001669b948b146250114f3e06
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505253"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a><span data-ttu-id="24521-103">将内置 Office 按钮集成到自定义控件组和选项卡中</span><span class="sxs-lookup"><span data-stu-id="24521-103">Integrate built-in Office buttons into custom control groups and tabs</span></span>

<span data-ttu-id="24521-104">可以使用加载项清单中的标记将内置 Office 按钮插入 Office 功能区上的自定义控件组中。</span><span class="sxs-lookup"><span data-stu-id="24521-104">You can insert built-in Office buttons into your custom control groups on the Office ribbon by using markup in the add-in's manifest.</span></span> <span data-ttu-id="24521-105"> (无法将自定义外接程序命令插入内置 Office 组。) 还可以将整个内置 Office 控件组插入到自定义功能区选项卡中。</span><span class="sxs-lookup"><span data-stu-id="24521-105">(You can't insert your custom add-in commands into a built-in Office group.) You can also insert entire built-in Office control groups into your custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="24521-106">本文假定您熟悉外接程序 [命令的基本概念一文](add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="24521-106">This article assumes that you are familiar with the article [Basic concepts for add-in commands](add-in-commands.md).</span></span> <span data-ttu-id="24521-107">如果你最近没有这样做，请查看它。</span><span class="sxs-lookup"><span data-stu-id="24521-107">Please review it if you haven't done so recently.</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="24521-108">本文中介绍的外接程序功能及标记 *仅在 PowerPoint 网页中可用*。</span><span class="sxs-lookup"><span data-stu-id="24521-108">The add-in feature and markup described in this article is *only available in PowerPoint on the web*.</span></span>
> - <span data-ttu-id="24521-109">本文中介绍的标记仅适用于支持要求集 **AddinCommands 1.3 的平台**。</span><span class="sxs-lookup"><span data-stu-id="24521-109">The markup described in this article only works on platforms that support requirement set **AddinCommands 1.3**.</span></span> <span data-ttu-id="24521-110">请参阅后一节 [不受支持的平台的行为](#behavior-on-unsupported-platforms)。</span><span class="sxs-lookup"><span data-stu-id="24521-110">See the later section [Behavior on unsupported platforms](#behavior-on-unsupported-platforms).</span></span>

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a><span data-ttu-id="24521-111">将内置控件组插入自定义选项卡</span><span class="sxs-lookup"><span data-stu-id="24521-111">Insert a built-in control group into a custom tab</span></span>

<span data-ttu-id="24521-112">若要将内置的 Office 控件组插入选项卡，请将 [OfficeGroup](../reference/manifest/customtab.md#officegroup) 元素添加为父元素中的子 `<CustomTab>` 元素。</span><span class="sxs-lookup"><span data-stu-id="24521-112">To insert a built-in Office control group into a tab, add an [OfficeGroup](../reference/manifest/customtab.md#officegroup) element as a child element in the parent `<CustomTab>` element.</span></span> <span data-ttu-id="24521-113">元素 `id` 的属性设置为内置组的 `<OfficeGroup>` ID。</span><span class="sxs-lookup"><span data-stu-id="24521-113">The `id` attribute of the of the `<OfficeGroup>` element is set to the ID of the built-in group.</span></span> <span data-ttu-id="24521-114">请参阅["查找控件和控件组的 ID"。](#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="24521-114">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="24521-115">以下标记示例将 Office Paragraph 控件组添加到自定义选项卡，并将它定位到自定义组之后。</span><span class="sxs-lookup"><span data-stu-id="24521-115">The following markup example adds the Office Paragraph control group to a custom tab and positions it to appear just after a custom group.</span></span>

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a><span data-ttu-id="24521-116">将内置控件插入自定义组</span><span class="sxs-lookup"><span data-stu-id="24521-116">Insert a built-in control into a custom group</span></span>

<span data-ttu-id="24521-117">若要将内置 Office 控件插入自定义组，请将 [OfficeControl](../reference/manifest/group.md#officecontrol) 元素添加为父元素中的子 `<Group>` 元素。</span><span class="sxs-lookup"><span data-stu-id="24521-117">To insert a built-in Office control into a custom group, add an [OfficeControl](../reference/manifest/group.md#officecontrol) element as a child element in the parent `<Group>` element.</span></span> <span data-ttu-id="24521-118">元素 `id` 的属性 `<OfficeControl>` 设置为内置控件的 ID。</span><span class="sxs-lookup"><span data-stu-id="24521-118">The `id` attribute of the `<OfficeControl>` element is set to the ID of the built-in control.</span></span> <span data-ttu-id="24521-119">请参阅["查找控件和控件组的 ID"。](#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="24521-119">See [Find the IDs of controls and control groups](#find-the-ids-of-controls-and-control-groups).</span></span>

<span data-ttu-id="24521-120">以下标记示例将 Office 上标控件添加到自定义组，并将它定位到自定义按钮的正后。</span><span class="sxs-lookup"><span data-stu-id="24521-120">The following markup example adds the Office Superscript control to a custom group and positions it to appear just after a custom button.</span></span>

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
> <span data-ttu-id="24521-121">用户可以在 Office 应用程序中自定义功能区。</span><span class="sxs-lookup"><span data-stu-id="24521-121">Users can customize the ribbon in the Office application.</span></span> <span data-ttu-id="24521-122">任何用户自定义项都将覆盖清单设置。</span><span class="sxs-lookup"><span data-stu-id="24521-122">Any user customizations will override your manifest settings.</span></span> <span data-ttu-id="24521-123">例如，用户可以从任何组中删除按钮，并从选项卡中删除任何组。</span><span class="sxs-lookup"><span data-stu-id="24521-123">For example, a user can remove a button from any group and remove any group from a tab.</span></span>

## <a name="find-the-ids-of-controls-and-control-groups"></a><span data-ttu-id="24521-124">查找控件和控件组的 ID</span><span class="sxs-lookup"><span data-stu-id="24521-124">Find the IDs of controls and control groups</span></span>

<span data-ttu-id="24521-125">支持的控件和控件组的 ID 在存储库 [Office 控件的文件中](https://github.com/OfficeDev/office-control-ids)。</span><span class="sxs-lookup"><span data-stu-id="24521-125">The IDs for supported controls and control groups are in files in the repo [Office Control IDs](https://github.com/OfficeDev/office-control-ids).</span></span> <span data-ttu-id="24521-126">按照该存储库的 ReadMe 文件中的说明操作。</span><span class="sxs-lookup"><span data-stu-id="24521-126">Follow the instructions in the ReadMe file of that repo.</span></span>

## <a name="behavior-on-unsupported-platforms"></a><span data-ttu-id="24521-127">不受支持的平台上的行为</span><span class="sxs-lookup"><span data-stu-id="24521-127">Behavior on unsupported platforms</span></span>

<span data-ttu-id="24521-128">如果外接程序安装在不支持要求集 [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则本文中描述的标记将被忽略，并且内置 Office 控件/组将不会显示在自定义组/选项卡中。</span><span class="sxs-lookup"><span data-stu-id="24521-128">If your add-in is installed on a platform that doesn't support [requirement set AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md), then the markup described in this article is ignored and the built-in Office controls/groups will not appear in your custom groups/tabs.</span></span> <span data-ttu-id="24521-129">若要防止加载项安装在不支持标记的平台上，请添加对清单部分中的要求集 `<Requirements>` 的引用。</span><span class="sxs-lookup"><span data-stu-id="24521-129">To prevent your add-in from being installed on platforms that don't support the markup, add a reference to the requirement set in the `<Requirements>` section of the manifest.</span></span> <span data-ttu-id="24521-130">有关说明，请参阅 [清单中的"设置 Requirements"元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="24521-130">For instructions, see [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span> <span data-ttu-id="24521-131">或者，可以将外接程序设计成在 **不支持 AddinCommands 1.3** 时具有备用体验，如 [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)代码中的"使用运行时检查"中所述。</span><span class="sxs-lookup"><span data-stu-id="24521-131">Alternatively, you can design your add-in to have an alternate experience when **AddinCommands 1.3** is not supported, as described in [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="24521-132">例如，如果您的外接程序包含假定内置按钮在自定义组中的说明，则您可能具有一个备用版本，该版本假定内置按钮仅在其常用位置。</span><span class="sxs-lookup"><span data-stu-id="24521-132">For example, if your add-in contains instructions that assume the built-in buttons are in your custom groups, you could have an alternate version that assumes that the built-in buttons are only in their usual places.</span></span>
