---
title: 清单文件中的 CustomTab 元素
description: 在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。
ms.date: 09/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: f8cdcd2c1a1e567f36d9d146ed4806b13d400dfe
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149581"
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
|  [InsertAfter](#insertafter)      | 否 |  指定自定义选项卡应紧接在指定的内置选项卡之后。Office：仅在 PowerPoint。  |
|  [InsertBefore](#insertbefore)      | 否 |  指定自定义选项卡应紧接在指定的内置选项卡Office之前。重要 **说明：仅在** PowerPoint。 |

### <a name="group"></a>组

可选，但如果不存在，则必须至少有一 **个 OfficeGroup** 元素。 查看 [Group 元素](group.md)。 清单中 **Group** 和 **OfficeGroup** 的顺序应为您希望它们显示在自定义选项卡上的顺序。如果有多个元素，则它们可以同时存在，但所有元素都必须在 **Label 元素** 之上。

### <a name="officegroup"></a>OfficeGroup

可选，但如果不存在，则必须至少有一 **个 Group** 元素。 代表内置控件Office组。 **id** 属性指定内置组Office ID。 若要查找内置组的 ID，请参阅查找控件和[控件组的 ID。](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) 清单中 **Group** 和 **OfficeGroup** 的顺序应为您希望它们显示在自定义选项卡上的顺序。如果有多个元素，则它们可以同时存在，但所有元素都必须在 **Label 元素** 之上。

> [!IMPORTANT]
> `OfficeGroup`元素在 Outlook 中不可用。

### <a name="label-tab"></a>标签（选项卡）

必需。 自定义选项卡的标签。**resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md)元素）中 **String** 元素的 **id** 属性的值。

### <a name="insertafter"></a>InsertAfter

可选。 指定自定义选项卡应紧接在指定的内置选项卡之后Office选项卡。元素的值为内置选项卡的 ID，如"TabHome"或"TabReview"。  (Find [the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present， must be after the **Label** element. 不能同时具有 **InsertAfter 和** **InsertBefore**。

> [!IMPORTANT]
> `InsertAfter`元素仅在 PowerPoint。

### <a name="insertbefore"></a>InsertBefore

可选。 指定自定义选项卡应紧接在指定的内置选项卡之前Office选项卡。元素的值为内置选项卡的 ID，如"TabHome"或"TabReview"。  (Find [the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present， must be after the **Label** element. 不能同时具有 **InsertAfter 和** **InsertBefore**。

> [!IMPORTANT]
> `InsertBefore`元素仅在 PowerPoint。
