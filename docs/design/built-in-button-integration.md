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
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs-preview"></a>将内置 Office 按钮集成到自定义控件组和选项卡中 (预览) 

您可以使用外接程序清单中的标记将内置 Office 按钮插入 Office 功能区上的自定义控件组中。  (无法将自定义外接程序命令插入内置 Office 组。 ) 也可以将整个内置 Office 控件组插入到自定义功能区选项卡中。

> [!NOTE]
> 本文假定您熟悉文章 [外接程序命令的基本概念](add-in-commands.md)。 如果你最近未执行此操作，请对其进行检查。

> [!IMPORTANT]
>
> - 本文中介绍的加载项功能和标记位于预览中， *仅适用于 web 上的 PowerPoint*。 我们建议您仅在测试和开发环境中尝试标记。 请勿在生产环境中或在业务关键型文档中使用预览标记。
> - 本文中所述的标记仅适用于支持要求集 **addincommand 1.3** 的平台。 有关 [不受支持的平台](#behavior-on-unsupported-platforms)，请参阅后续章节行为。

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>在自定义选项卡中插入内置控件组

若要在选项卡中插入内置 Office 控件组，请在 parent 元素中将 [OfficeGroup](../reference/manifest/customtab.md#officegroup) 元素添加为子元素 `<CustomTab>` 。 将 `id` 元素的属性 `<OfficeGroup>` 设置为内置组的 ID。 请参阅 [查找控件和控件组的 id](#find-the-ids-of-controls-and-control-groups)。

下面的标记示例将 "Office 段落" 控件组添加到自定义选项卡，并将其放置在自定义组的紧后面。

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a>将内置控件插入到自定义组中

若要将内置 Office 控件插入到自定义组中，请在父元素中将 [OfficeControl](../reference/manifest/group.md#officecontrol) 元素添加为子元素 `<Group>` 。 `id`元素的属性 `<OfficeControl>` 设置为内置控件的 ID。 请参阅 [查找控件和控件组的 id](#find-the-ids-of-controls-and-control-groups)。

下面的标记示例将 Office 上标控件添加到自定义组，并将其放置在自定义按钮的紧后面。

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
> 用户可以在 Office 应用程序中自定义功能区。 任何用户自定义设置都将覆盖您的清单设置。 例如，用户可以从任何组中删除按钮，并从选项卡中删除任何组。

## <a name="find-the-ids-of-controls-and-control-groups"></a>查找控件和控件组的 Id

支持的控件和控件组的 Id 位于存储库 [Office 控件 id](https://github.com/OfficeDev/office-control-ids)的文件中。 按照该存储库的自述文件中的说明进行操作。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果外接程序安装在不支持 [要求集 addincommand 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则会忽略本文中所述的标记，并且内置的 Office 控件/组将不会显示在自定义组/选项卡中。 若要防止外接程序安装在不支持标记的平台上，请在清单的部分中添加对要求集的引用 `<Requirements>` 。 有关说明，请参阅 [在清单中设置需求元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。 此外，还可以设计外接程序，使其在 **addincommand 1.3** 不受支持时具有备用体验，如 [JavaScript 代码中的使用运行时检查](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)中所述。 例如，如果您的外接程序包含假设内置按钮位于您的自定义组中的说明，则可以使用一个替代版本，它假定内置按钮仅位于其通常的位置。
