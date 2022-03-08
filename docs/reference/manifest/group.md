---
title: 清单文件中 Group 元素
description: 在选项卡中定义一组 UI 控件。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: b4717f6aeff3cd8ac34ee289252054417c489b89
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340461"
---
# <a name="group-element"></a>Group 元素

在选项卡中定义一组 UI 控件。在自定义选项卡上，外接程序可以创建多个组。 外接程序限定到一个自定义选项卡。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) 

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  是  | 组的唯一 ID。|

### <a name="id-attribute"></a>id attribute

必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 这必须在清单中所有 Group 元素中是唯一的。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Label](#label)      | 是 |  组的标签。  |
|  [Icon](icon.md)      | 是 |  组的图像。 在加载项Outlook不支持。 |
|  [Control](#control)    | 否 |  代表 Control 对象。 可以是零个或多个。  |
|  [OfficeControl](#officecontrol)  | 否 | 表示内置控件之Office控件。 可以是零个或多个。 在加载项Outlook不支持。|
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | 否 |  指定组是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。 在加载项Outlook不支持。 |

### <a name="label"></a>标签

必需。 组的标签。 **resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。

### <a name="icon"></a>Icon

必需。 如果选项卡包含大量组，并且程序窗口已调整大小，则可能会改为显示指定的图像。

> [!NOTE]
> 此子元素在加载项Outlook支持。

### <a name="control"></a>控制

可选，但如果不存在，则必须至少有一 **个 OfficeControl**。 有关受支持的控件类型的详细信息，请参阅 [Control](control.md) 元素。 清单中的 **Control** 和 **OfficeControl** 顺序是可互换的，如果有多个元素，则它们可以相互组成，但所有元素都必须位于 **Icon** 元素下方。

```xml
<Group id="Contoso.CustomTab1.group1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Contoso.Button1">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a>OfficeControl

可选，但如果不存在，则必须至少有一个 **Control**。 在包含元素的组中Office一个或多个内置控件`<OfficeControl>`。 属性`id`指定内置控件Office ID。 若要查找控件的 ID，请参阅查找控件 [和控件组的 ID](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 清单中的 **Control** 和 **OfficeControl** 顺序是可互换的，如果有多个元素，则它们可以相互组成，但所有元素都必须位于 **Icon** 元素下方。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

> [!NOTE]
> 此子元素在加载项Outlook支持。

```xml
<Group id="Contoso.CustomTab2.group2">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Contoso.Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

可选 (布尔值) 。 指定是否在支持 API  的应用程序和平台组合上隐藏组，该 API 在运行时在功能区上安装自定义上下文选项卡。 默认值（如果不存在）为 `false`。 如果使用， **则 OverriddenByRibbonApi** 必须是 *Group 的第一* 个 **子级**。 有关详细信息，请参阅 [OverriddenByRibbonApi](overriddenbyribbonapi.md)。

> [!NOTE]
> 此子元素在加载项Outlook支持。

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.CustomTab3">
    <Group id="Contoso.CustomTab3.group1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
