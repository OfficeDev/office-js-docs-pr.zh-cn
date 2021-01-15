---
title: 清单文件中的 CustomTab 元素
description: 在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 642222af02431814e4e64141504911c67ca829fa
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771324"
---
# <a name="customtab-element"></a>CustomTab 元素

在功能区上，指定外接程序命令的选项卡和组。 这可能位于默认选项卡（“主页”、“邮件”或“会议”）上，或位于外接程序定义的自定义选项卡上。

在自定义选项卡上，加载项可以具有自定义组或内置组。 外接程序限定到一个自定义选项卡。

**id** 属性在清单中必须是唯一的。

> [!IMPORTANT]
> 在 Mac 上的 Outlook 中，该元素 `CustomTab` 不可用，因此您必须改为使用[OfficeTab。](officetab.md)

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | 否 |  定义一组命令。  |
|  [OfficeGroup](#officegroup)      | 否 |  代表内置的 Office 控件组。  |
|  [Label](#label-tab)      | 是 |  CustomTab 或组的标签。  |
|  [InsertAfter](#insertafter)      | 否 |  指定自定义选项卡应紧接在指定的内置 Office 选项卡之后。  |
|  [InsertBefore](#insertbefore)      | 否 |  指定自定义选项卡应紧接在指定的内置 Office 选项卡之前。  |

### <a name="group"></a>组

可选，但如果不存在，则必须至少有一 **个 OfficeGroup** 元素。 查看 [Group 元素](group.md)。 清单 **中组** 和 **OfficeGroup** 的顺序应该是您希望它们显示在自定义选项卡上的顺序。如果有多个元素，它们可能会同时存在，但所有元素都必须在 **Label 元素** 之上。

### <a name="officegroup"></a>OfficeGroup

可选，但如果不存在，则必须至少有一 **个 Group** 元素。 代表内置的 Office 控件组。 **id** 属性指定内置 Office 组的 ID。 若要查找内置组的 ID，请参阅"查找控件和[控件组的 ID"。](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) 清单 **中组** 和 **OfficeGroup** 的顺序应该是您希望它们显示在自定义选项卡上的顺序。如果有多个元素，它们可能会同时存在，但所有元素都必须在 **Label 元素** 之上。

### <a name="label-tab"></a>标签（选项卡）

必需。 自定义选项卡的标签。**resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)

### <a name="insertafter"></a>InsertAfter

可选。 指定自定义选项卡应紧接在指定的内置 Office 选项卡之后。元素的值是内置选项卡的 ID，如"TabHome"或"TabReview"。  ([查找控件和](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)控件组的标识。) 如果存在，则必须在 **Label 元素** 之后。 不能同时具有 **InsertAfter 和** **InsertBefore。**

### <a name="insertbefore"></a>InsertBefore

可选。 指定自定义选项卡应紧接在指定的内置 Office 选项卡之前。元素的值是内置选项卡的 ID，如"TabHome"或"TabReview"。  ([查找控件和](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)控件组的标识。) 如果存在，则必须在 **Label 元素** 之后。 不能同时具有 **InsertAfter 和** **InsertBefore。**

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
