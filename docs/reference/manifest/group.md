---
title: 清单文件中 Group 元素
description: 在选项卡中定义一组 UI 控件。
ms.date: 06/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 09260ab52910235ab63149769cc989ffbda03ffb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152621"
---
# <a name="group-element"></a>Group 元素

在选项卡中定义一组 UI 控件。在自定义选项卡上，外接程序可以创建多个组。 外接程序限定到一个自定义选项卡。

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
|  [Icon](icon.md)      | 是 |  组的图像。 在加载项Outlook不支持。 |
|  [Control](#control)    | 否 |  代表 Control 对象。 可以是零个或多个。  |
|  [OfficeControl](#officecontrol)  | 否 | 代表内置控件之Office控件。 可以是零个或多个。 在加载项Outlook不支持。|
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | 否 |  指定组是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。 在加载项Outlook不支持。 |

### <a name="label"></a>标签

必需。 组的标签。 **resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md)元素）中 **String** 元素的 **id** 属性的值。

### <a name="icon"></a>Icon

必需。 如果选项卡包含大量组，并且程序窗口已调整大小，则可能会改为显示指定的图像。

> [!NOTE]
> 此子元素在加载项中Outlook支持。

### <a name="control"></a>控件

可选，但如果不存在，则必须至少有一 **个 OfficeControl**。 有关受支持的控件类型的详细信息，请参阅 [Control](control.md) 元素。 清单中的 **Control** 和 **OfficeControl** 顺序是可互换的，如果有多个元素，则它们可以相互组成，但所有元素都必须位于 **Icon** 元素下方。

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

可选，但如果不存在，则必须至少有一个 **Control**。 在包含元素的组中Office一个或多个内置控件 `<OfficeControl>` 。 `id`属性指定内置控件Office ID。 若要查找控件的 ID，请参阅查找控件[和控件组的 ID。](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) 清单中的 **Control** 和 **OfficeControl** 顺序是可互换的，如果有多个元素，则它们可以相互组成，但所有元素都必须位于 **Icon** 元素下方。

> [!NOTE]
> 此子元素在加载项中Outlook支持。

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

可选 (布尔) 。 指定是否在支持API 的应用程序和平台组合上隐藏组，该 API 在运行时在功能区上安装自定义上下文选项卡。 默认值（如果不存在）为 `false` 。 如果使用 **，OverriddenByRibbonApi** 必须是 *组* 的第一个 **子级**。 有关详细信息，请参阅 [OverriddenByRibbonApi](overriddenbyribbonapi.md)。

> [!NOTE]
> 此子元素在加载项中Outlook支持。

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
