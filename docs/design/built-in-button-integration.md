---
title: 将内置 Office 按钮集成到自定义控件组和选项卡中
description: 了解如何在自定义命令组和 Office 功能区选项卡中包含内置 Office 按钮。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dc706fcd0b049647847a73f7c40144dba9df0e2
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659785"
---
# <a name="integrate-built-in-office-buttons-into-custom-control-groups-and-tabs"></a>将内置 Office 按钮集成到自定义控件组和选项卡中

可以使用加载项清单中的标记将内置 Office 按钮插入 Office 功能区上的自定义控制组中。  (不能将自定义外接程序命令插入内置 Office 组。) 还可以将整个内置 Office 控件组插入自定义功能区选项卡中。

> [!NOTE]
> 本文假定你熟悉 [加载项命令的基本概念](add-in-commands.md)一文。 如果你最近没有这样做，请查看它。

> [!IMPORTANT]
>
> - 本文中所述的加载项功能和标记 *仅在PowerPoint web 版中可用*。
> - 本文中所述的标记仅适用于支持要求集 **AddinCommands 1.3** 的平台。 请参阅后面部分“ [不受支持的平台上的行为”](#behavior-on-unsupported-platforms)部分。

## <a name="insert-a-built-in-control-group-into-a-custom-tab"></a>将内置控件组插入自定义选项卡

若要在选项卡中插入内置的 Office 控件组，请在父 **\<CustomTab\>** 元素中添加 [OfficeGroup](/javascript/api/manifest/customtab#officegroup) 元素作为子元素。 元素 `id` 的 **\<OfficeGroup\>** 属性设置为内置组的 ID。 请参阅 [“查找控件和控件组的 ID](#find-the-ids-of-controls-and-control-groups)”。

以下标记示例将 Office 段落控制组添加到自定义选项卡，并将其定位为在自定义组之后显示。

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

若要将内置 Office 控件插入自定义组，请在父 **\<Group\>** 元素中添加 [OfficeControl](/javascript/api/manifest/group#officecontrol) 元素作为子元素。 元素 `id` 的 **\<OfficeControl\>** 属性设置为内置控件的 ID。 请参阅 [“查找控件和控件组的 ID](#find-the-ids-of-controls-and-control-groups)”。

以下标记示例将 Office 上标控件添加到自定义组，并将其定位为在自定义按钮之后显示。

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
> 用户可以在 Office 应用程序中自定义功能区。 任何用户自定义项都将覆盖清单设置。 例如，用户可以从任何组中删除按钮，并从选项卡中删除任何组。

## <a name="find-the-ids-of-controls-and-control-groups"></a>查找控件和控件组的 ID

支持的控件和控件组的 ID 位于存储库 [Office 控件 ID](https://github.com/OfficeDev/office-control-ids) 中的文件中。 按照该存储库的自述文件中的说明操作。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果外接程序安装在不支持 [要求集 AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets) 的平台上，则忽略本文中所述的标记，并且内置的 Office 控件/组不会显示在自定义组/选项卡中。 若要防止在不支持标记的平台上安装加载项，请在清单部分中 **\<Requirements\>** 添加对要求集的引用。 有关说明，请参阅 [指定哪些 Office 版本和平台可以托管外接程序](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。 或者，在不支持 **AddinCommands 1.3** 时设计外接程序以获得体验，如“设计”中所述 [，以获得备用体验](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences)。 例如，如果外接程序包含的说明假定内置按钮位于自定义组中，则可以设计一个版本，假定内置按钮仅位于其常用位置。
