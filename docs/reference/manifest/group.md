---
title: 清单文件中 Group 元素
description: 在选项卡中定义一组 UI 控件。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173960"
---
# <a name="group-element"></a>Group 元素

在选项卡中定义一组 UI 控件。在自定义选项卡上，加载项可以创建多个组。 外接程序限定到一个自定义选项卡。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  是  | 组的唯一 ID。|

### <a name="id-attribute"></a>id attribute

必需。 组的唯一标识符。 是一个最多为 125 个字符的字符串。 该字符串在清单内必须是唯一的，否则组将不能呈现。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Label](#label)      | 是 |  CustomTab 或组的标签。  |
|  [Icon](icon.md)      | 是 |  组的图像。  |
|  [Control](#control)    | 否 |  代表一个 Control 对象。 可以是零个或多个。  |
|  [OfficeControl](#officecontrol)  | 否 | 表示一个内置的 Office 控件。 可以是零个或多个。 |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | 否 |  指定组是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。  |

### <a name="label"></a>标签

必需。 组的标签。 **resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)

### <a name="icon"></a>Icon

必需。 如果选项卡包含大量组，并且程序窗口调整了大小，则可能会改为显示指定的图像。

### <a name="control"></a>控制

可选，但如果不存在，则必须至少有一 **个 OfficeControl**。 有关支持的控件类型的详细信息，请参阅 [Control](control.md) 元素。 清单中的 **Control** 和 **OfficeControl** 顺序是可互换的，如果有多个元素，它们可以相互交集，但所有元素都必须位于 **Icon** 元素下方。

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a>OfficeControl

可选，但如果不存在，则必须至少有一个 **控件**。 在包含元素的组中包括一个或多个内置 Office `<OfficeControl>` 控件。 `id`该属性指定内置 Office 控件的 ID。 若要查找控件的 ID，请参阅"[查找控件和控件组的 ID"。](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) 清单中的 **Control** 和 **OfficeControl** 顺序是可互换的，如果有多个元素，它们可以相互交集，但所有元素都必须位于 **Icon** 元素下方。

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

可选 (布尔) 。 指定是否在支持API 的应用程序和平台组合上隐藏组，该 API 在运行时在功能区上安装自定义上下文选项卡。 默认值（如果不存在）为 `false` 。 如果使用 **，OverriddenByRibbonApi** 必须是 *组* 的第一 **个子级**。 有关详细信息，请参阅 [OverriddenByRibbonApi](overriddenbyribbonapi.md)。

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
