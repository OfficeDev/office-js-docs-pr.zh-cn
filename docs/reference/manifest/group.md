---
title: 清单文件中的 Group 元素
description: 定义选项卡中的一组 UI 控件。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 6ee8d499767eccb95b4fdf9ceb91dd2cd12bce95
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087943"
---
# <a name="group-element"></a>Group 元素

定义选项卡中的一组 UI 控件。在自定义选项卡上，加载项可以创建多个组。 外接程序限定到一个自定义选项卡。

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
|  [Control](#control)    | 否 |  表示控件对象。 可以是零个或多个。  |
|  [OfficeControl](#officecontrol)  | 否 | 表示内置 Office 控件之一。 可以是零个或多个。 |

### <a name="label"></a>标签

必需。 组的标签。 **Resid** 属性必须设置为 [Resources](resources.md)元素中的 **ShortStrings** 元素中 **String** 元素的 **id** 属性的值。

### <a name="icon"></a>Icon

必需。 如果某个选项卡包含大量组，并且该程序窗口已调整大小，则会改为显示指定的图像。

### <a name="control"></a>控制

可选，但如果不存在，则必须至少有一个 **OfficeControl**。 有关受支持的控件类型的详细信息，请参阅 [Control](control.md) 元素。 在清单中， **Control** 和 **OfficeControl** 的顺序是可互换的，如果存在多个元素，则可以是混合的，但所有元素都必须位于 **Icon** 元素的下面。

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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

可选，但如果不存在，则必须至少有一个 **控件**。 在带元素的组中包含一个或多个内置 Office 控件 `<OfficeControl>` 。 `id`属性指定内置 Office 控件的 ID。 若要查找控件的 ID，请参阅 [查找控件和控件组的 id](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 在清单中， **Control** 和 **OfficeControl** 的顺序是可互换的，如果存在多个元素，则可以是混合的，但所有元素都必须位于 **Icon** 元素的下面。

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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
