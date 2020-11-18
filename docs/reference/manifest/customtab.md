---
title: 清单文件中的 CustomTab 元素
description: 在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 99670b27d963060a008899a8808ca967cfd710a6
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087936"
---
# <a name="customtab-element"></a>CustomTab 元素

在功能区上，为您的外接程序命令指定选项卡和组。 这可能位于默认选项卡（“主页”、“邮件”或“会议”）上，或位于外接程序定义的自定义选项卡上。

在自定义选项卡上，加载项可以具有自定义或内置组。 外接程序限定到一个自定义选项卡。

**Id** 属性在清单中必须是唯一的。

> [!IMPORTANT]
> 在 Mac 上的 Outlook 中，该元素不可用， `CustomTab` 因此您必须改用 [OfficeTab](officetab.md) 。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | 否 |  定义一组命令。  |
|  [OfficeGroup](#officegroup)      | 否 |  代表内置 Office 控件组。  |
|  [Label](#label-tab)      | 是 |  CustomTab 或组的标签。  |
|  [InsertAfter](#insertafter)      | 否 |  指定自定义选项卡应紧跟在指定的内置 "Office" 选项卡之后。  |
|  [InsertBefore](#insertbefore)      | 否 |  指定自定义选项卡应紧跟在指定的内置 "Office" 选项卡之前。  |

### <a name="group"></a>Group

可选，但如果不存在，则必须至少有一个 **OfficeGroup** 元素。 查看 [Group 元素](group.md)。 清单和 **OfficeGroup** 在清单 **中的顺序** 应是您希望它们出现在 "自定义" 选项卡上的顺序。如果有多个元素，则可以是混合的，但所有元素都必须位于 **Label** 元素的上方。

### <a name="officegroup"></a>OfficeGroup

可选，但如果不存在，则必须至少有一个 **Group** 元素。 代表内置 Office 控件组。 **Id** 属性指定内置 Office 组的 id。 若要查找内置组的 ID，请参阅 [查找控件和控件组的 id](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 清单和 **OfficeGroup** 在清单 **中的顺序** 应是您希望它们出现在 "自定义" 选项卡上的顺序。如果有多个元素，则可以是混合的，但所有元素都必须位于 **Label** 元素的上方。

### <a name="label-tab"></a>标签（选项卡）

必需。 自定义选项卡的标签。**Resid** 属性必须设置为 [Resources](resources.md)元素中的 **ShortStrings** 元素中 **String** 元素的 **id** 属性的值。

### <a name="insertafter"></a>InsertAfter

可选。 指定自定义选项卡应紧跟在指定的内置 "Office" 选项卡之后。元素的值是内置选项卡的 ID，如 "TabHome" 或 "TabReview"。  (请参阅 [查找控件和控件组的 id](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 ) 如果存在，则必须位于 **Label** 元素之后。 您不能同时具有 **InsertAfter** 和 **InsertBefore**。

### <a name="insertbefore"></a>InsertBefore

可选。 指定自定义选项卡应紧跟在指定的内置 "Office" 选项卡之前。元素的值是内置选项卡的 ID，如 "TabHome" 或 "TabReview"。  (请参阅 [查找控件和控件组的 id](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 ) 如果存在，则必须位于 **Label** 元素之后。 您不能同时具有 **InsertAfter** 和 **InsertBefore**。

## <a name="customtab-example"></a>CustomTab 示例

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
