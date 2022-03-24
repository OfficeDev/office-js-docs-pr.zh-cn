---
title: 将内置控件Office集成到自定义控件组和选项卡中
description: 了解如何在自定义命令组Office自定义命令组和自定义功能区上的选项卡Office按钮。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: b9f334bdc84353409c81059a3f5cfd60bbb4c0fa
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743087"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>将内置控件Office集成到自定义控件组和选项卡中

可以使用加载项清单中的Office，将内置控件按钮插入到 Office 功能区上的自定义控件组中。  (无法将自定义外接程序命令插入内置 Office 组。) 还可以将整个内置 Office 控件组插入自定义功能区选项卡。

> [!NOTE]
> 本文假定您熟悉文章 [Basic concepts for add-in commands](add-in-commands.md)。 如果你最近未这样做，请查看它。

> [!IMPORTANT]
>
> - 本文中介绍的加载项功能与 *标记仅在 PowerPoint web 版*。
> - 本文中介绍的标记仅适用于支持要求集 **AddinCommands 1.3 的平台**。 请参阅下一节 [不受支持的平台上的行为](#behavior-on-unsupported-platforms)。

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>将内置控件组插入自定义选项卡

若要将内置控件Office插入选项卡，请将 [OfficeGroup](../reference/manifest/customtab.md#officegroup) 元素添加为 **父 CustomTab** 元素中的子元素。 `id` **OfficeGroup** 元素的 属性设置为内置组的 ID。 请参阅 [查找控件和控件组的 ID](#find-the-ids-of-controls-and-control-groups)。

以下标记示例将 Office Paragraph 控件组添加到自定义选项卡，并将它定位到自定义组之后。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

## <a name="insert-a-built-in-control-into-a-custom-group"></a>将内置控件插入自定义组

若要将内置控件Office自定义组中，请将 [OfficeControl](../reference/manifest/group.md#officecontrol) 元素添加为 **父 Group** 元素中的子元素。 `id` **OfficeControl** 元素的 属性设置为内置控件的 ID。 请参阅 [查找控件和控件组的 ID](#find-the-ids-of-controls-and-control-groups)。

以下标记示例将上标Office添加到自定义组，并将它定位到自定义按钮之后。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group1">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button1">
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
> 用户可以在应用程序应用程序中自定义Office功能区。 任何用户自定义设置都将覆盖清单设置。 例如，用户可以从任何组中删除按钮，并从选项卡中删除任何组。

## <a name="find-the-ids-of-controls-and-control-groups"></a>查找控件和控件组的 ID

支持的控件和控件组的 ID 在控件OFFICE[文件中](https://github.com/OfficeDev/office-control-ids)。 按照该存储库的 ReadMe 文件中的说明操作。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果外接程序安装在不支持要求集 [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md) 的平台上，则本文中描述的标记将被忽略，并且内置 Office 控件/组将不会显示在自定义组/选项卡中。 若要防止外接程序安装在不支持标记的平台上，请添加对清单的"要求"部分的要求集的引用。 有关说明，请参阅[指定Office哪些版本和平台可以托管你的外接程序](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。 或者，设计外接程序以在 **AddinCommands 1.3** 不受支持时获得体验，如设计 [备用体验中所述](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences)。 例如，如果您的外接程序包含假定内置按钮在自定义组中的说明，您可以设计一个版本，假定内置按钮仅在其常用位置。
