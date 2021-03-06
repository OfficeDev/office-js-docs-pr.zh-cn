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
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>将内置 Office 按钮集成到自定义控件组和选项卡中

可以使用加载项清单中的标记将内置 Office 按钮插入 Office 功能区上的自定义控件组中。  (无法将自定义外接程序命令插入内置 Office 组。) 还可以将整个内置 Office 控件组插入到自定义功能区选项卡中。

> [!NOTE]
> 本文假定您熟悉外接程序 [命令的基本概念一文](add-in-commands.md)。 如果你最近没有这样做，请查看它。

> [!IMPORTANT]
>
> - 本文中介绍的外接程序功能及标记 *仅在 PowerPoint 网页中可用*。
> - 本文中介绍的标记仅适用于支持要求集 **AddinCommands 1.3 的平台**。 请参阅后一节 [不受支持的平台的行为](#behavior-on-unsupported-platforms)。

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>将内置控件组插入自定义选项卡

若要将内置的 Office 控件组插入选项卡，请将 [OfficeGroup](../reference/manifest/customtab.md#officegroup) 元素添加为父元素中的子 `<CustomTab>` 元素。 元素 `id` 的属性设置为内置组的 `<OfficeGroup>` ID。 请参阅["查找控件和控件组的 ID"。](#find-the-ids-of-controls-and-control-groups)

以下标记示例将 Office Paragraph 控件组添加到自定义选项卡，并将它定位到自定义组之后。

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

## <a name="insert-a-built-in-control-into-a-custom-group"></a>将内置控件插入自定义组

若要将内置 Office 控件插入自定义组，请将 [OfficeControl](../reference/manifest/group.md#officecontrol) 元素添加为父元素中的子 `<Group>` 元素。 元素 `id` 的属性 `<OfficeControl>` 设置为内置控件的 ID。 请参阅["查找控件和控件组的 ID"。](#find-the-ids-of-controls-and-control-groups)

以下标记示例将 Office 上标控件添加到自定义组，并将它定位到自定义按钮的正后。

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
> 用户可以在 Office 应用程序中自定义功能区。 任何用户自定义项都将覆盖清单设置。 例如，用户可以从任何组中删除按钮，并从选项卡中删除任何组。

## <a name="find-the-ids-of-controls-and-control-groups"></a>查找控件和控件组的 ID

支持的控件和控件组的 ID 在存储库 [Office 控件的文件中](https://github.com/OfficeDev/office-control-ids)。 按照该存储库的 ReadMe 文件中的说明操作。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果外接程序安装在不支持要求集 [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则本文中描述的标记将被忽略，并且内置 Office 控件/组将不会显示在自定义组/选项卡中。 若要防止加载项安装在不支持标记的平台上，请添加对清单部分中的要求集 `<Requirements>` 的引用。 有关说明，请参阅 [清单中的"设置 Requirements"元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。 或者，可以将外接程序设计成在 **不支持 AddinCommands 1.3** 时具有备用体验，如 [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)代码中的"使用运行时检查"中所述。 例如，如果您的外接程序包含假定内置按钮在自定义组中的说明，则您可能具有一个备用版本，该版本假定内置按钮仅在其常用位置。
