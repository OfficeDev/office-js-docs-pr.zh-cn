---
title: 清单文件中的 CustomTab 元素
description: 在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: de6233966abea4de423f255bda3c9e6e38ff5037c760c90cae7c8a1c7ca6ab2e
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085049"
---
# <a name="customtab-element"></a>CustomTab 元素

在功能区上，指定外接程序命令的选项卡和组。 这可能位于默认选项卡（“主页”、“邮件”或“会议”）上，或位于外接程序定义的自定义选项卡上。

在自定义选项卡上，外接程序可以具有自定义组或内置组。 外接程序限定到一个自定义选项卡。

**id** 属性在清单中必须是唯一的。

> [!IMPORTANT]
> 在 Outlook Mac 上，元素不可用 `CustomTab` ，因此您必须改为使用[OfficeTab。](officetab.md)

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | 否 |  定义一组命令。  |
|  [OfficeGroup](#officegroup)      | 否 |  代表内置控件Office组。 **重要** 提示：在Outlook。 |
|  [Label](#label-tab)      | 是 |  CustomTab 或组的标签。  |
|  [InsertAfter](#insertafter)      | 否 |  指定自定义选项卡应紧接在指定的内置选项卡之后。Office：在自定义选项卡Outlook。  |
|  [InsertBefore](#insertbefore)      | 否 |  指定自定义选项卡应紧接在指定的内置选项卡之前。Office：在自定义选项卡Outlook。  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | 否 |  指定自定义选项卡是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。 **重要** 提示：在Outlook。 |

### <a name="group"></a>Group

可选，但如果不存在，则必须至少有一 **个 OfficeGroup** 元素。 查看 [Group 元素](group.md)。 清单中 **Group** 和 **OfficeGroup** 的顺序应该是您希望它们显示在自定义选项卡上的顺序。如果有多个元素，则它们可以同时存在，但所有元素都必须在 **Label 元素** 之上。

### <a name="officegroup"></a>OfficeGroup

可选，但如果不存在，则必须至少有一 **个 Group** 元素。 代表内置控件Office组。 **id** 属性指定内置组Office ID。 若要查找内置组的 ID，请参阅查找控件和[控件组的 ID。](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) 清单中 **Group** 和 **OfficeGroup** 的顺序应该是您希望它们显示在自定义选项卡上的顺序。如果有多个元素，则它们可以同时存在，但所有元素都必须在 **Label 元素** 之上。

> [!IMPORTANT]
> `OfficeGroup`元素在 Outlook 中不可用。

### <a name="label-tab"></a>标签（选项卡）

必填。 自定义选项卡的标签。**resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md)元素）中 **String** 元素的 **id** 属性的值。

### <a name="insertafter"></a>InsertAfter

可选。 指定自定义选项卡应紧接在指定的内置选项卡之后Office选项卡。元素的值为内置选项卡的 ID，如"TabHome"或"TabReview"。  (Find [the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present， must be after the **Label** element. 不能同时具有 **InsertAfter 和** **InsertBefore**。

> [!IMPORTANT]
> `InsertAfter`元素在 Outlook 中不可用。

### <a name="insertbefore"></a>InsertBefore

可选。 指定自定义选项卡应紧接在指定的内置选项卡Office之前。元素的值为内置选项卡的 ID，如"TabHome"或"TabReview"。  (Find [the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present， must be after the **Label** element. 不能同时具有 **InsertAfter 和** **InsertBefore**。

> [!IMPORTANT]
> `InsertBefore`元素在 Outlook 中不可用。

### <a name="overriddenbyribbonapi"></a>OverriddenByRibbonApi

可选 (布尔) 。 指定在支持 API 的应用程序和平台组合上是否隐藏 **CustomTab，** 该 API 在运行时在功能区上安装自定义上下文选项卡。 默认值（如果不存在）为 `false` 。 如果使用 **，OverriddenByRibbonApi** 必须是 **CustomTab 的第一个子级**。  有关详细信息，请参阅 [OverriddenByRibbonApi](overriddenbyribbonapi.md)。

> [!IMPORTANT]
> `OverriddenByRibbonApi`元素在 Outlook 中不可用。

## <a name="customtab-example"></a>CustomTab 示例

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
