---
title: 清单文件中的 CustomTab 元素
description: 在功能区上，可以为它们的外接程序命令指定使用哪种选项卡和组。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6a9540fd7e98464681a90021a36f7a7529186f7f
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340111"
---
# <a name="customtab-element"></a>CustomTab 元素

为"自定义"功能区定义Office选项卡。 将外接程序的功能区控件和组添加到内部版本 Office 选项卡或您自己的自定义选项卡。使用 **CustomTab** 元素将自定义选项卡添加到功能区。 在自定义选项卡上，外接程序可以具有自定义组或内置组。 外接程序限定到一个自定义选项卡。

> [!IMPORTANT]
> 在 Outlook Mac 上，**CustomTab** 元素不可用，但您可以改为将自定义控件组放在内置 [OfficeTabs](officetab.md) 之一上。 不能将 *内置组**置于任何* 平台上的内置Outlook选项卡上。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

> [!NOTE]
> 一些子元素无效邮件架构中。 请参阅 [子元素](#child-elements)。

**与以下要求集相关联**：

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)
- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md). 某些子元素需要。 请参阅 [子元素](#child-elements)。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  是  | 自定义选项卡的唯一 ID。|

### <a name="id-attribute"></a>id attribute

必需。 自定义选项卡的唯一标识符。它是一个最多包含 125 个字符的字符串。 这必须在清单中是唯一的。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Group](group.md)      | 否 |  定义一组命令。  |
|  [OfficeGroup](#officegroup)      | 否 |  代表内置控件Office组。 **重要** 提示：在 Outlook 中不可用。 |
|  [Label](#label-tab)      | 是 |  CustomTab 的标签。  |
|  [InsertAfter](#insertafter)      | 否 |  指定自定义选项卡应紧接在指定的内置选项卡Office。**重要** 说明：仅在 PowerPoint。 |
|  [InsertBefore](#insertbefore)      | 否 |  指定自定义选项卡应紧接在指定的内置选项卡Office之前。重要 **说明：仅在** PowerPoint。 |

### <a name="group"></a>Group

可选，但如果不存在，则必须至少有一 **个 OfficeGroup** 元素。 查看 [Group 元素](group.md)。 清单 **中 Group** 和 **OfficeGroup** 的顺序应为您希望它们显示在自定义选项卡上的顺序。如果有多个元素，则它们可以同时存在，但所有元素都必须在 **Label 元素** 之上。

### <a name="officegroup"></a>OfficeGroup

可选，但如果不存在，则必须至少有一 **个 Group** 元素。 代表内置控件Office组。 **id** 属性指定内置组Office ID。 若要查找内置组的 ID，请参阅查找控件和 [控件组的 ID](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)。 清单 **中 Group** 和 **OfficeGroup** 的顺序应为您希望它们显示在自定义选项卡上的顺序。如果有多个元素，则它们可以同时存在，但所有元素都必须在 **Label 元素** 之上。

> [!IMPORTANT]
> **OfficeGroup** 元素在 Outlook 中不可用。 在 PowerPoint 中，它在 Mac 和 Windows 预览版中;但对于 PowerPoint web 版 中的生产外接程序PowerPoint web 版。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="label-tab"></a>标签（选项卡）

必需项。 自定义选项卡的标签。**resid** 属性的长度不能超过 32 个字符，并且必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertafter"></a>InsertAfter

可选。 指定自定义选项卡应紧接在指定的内置选项卡之后Office选项卡。元素的值是内置选项卡的 ID，如 `TabHome` 或 `TabReview`。  有关内置选项卡的列表，请参阅 [OfficeTab](officetab.md)。 如果存在，则必须在 **Label 元素** 之后。 不能同时具有 **InsertAfter 和** **InsertBefore**。

> [!IMPORTANT]
> **InsertAfter** 元素仅在 PowerPoint 中可用。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)

### <a name="insertbefore"></a>InsertBefore

可选。 指定自定义选项卡应紧接在指定的内置选项卡之前Office选项卡。元素的值是内置选项卡的 ID，如 `TabHome` 或 `TabReview`。 元素的值是内置选项卡的 ID，如 `TabHome` 或 `TabReview`。  有关内置选项卡的列表，请参阅 [OfficeTab](officetab.md)。 如果存在，则必须在 **Label 元素** 之后。 不能同时具有 **InsertAfter 和** **InsertBefore**。

> [!IMPORTANT]
> **InsertBefore** 元素仅在 PowerPoint。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.3](../requirement-sets/add-in-commands-requirement-sets.md)


## <a name="examples"></a>示例

以下标记示例将 Office Paragraph 控件组添加到自定义选项卡，并将它定位到自定义组之后。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom1">
    <Group id="Contoso.TabCustom1.group1">
       <!-- additional markup omitted -->
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```

以下标记示例将上标Office添加到自定义组，并将它定位到自定义按钮之后。

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
        <Label resid="residCustomTabGroupLabel"/>
        <Icon>
            <bt:Image size="16" resid="blue-icon-16" />
            <bt:Image size="32" resid="blue-icon-32" />
            <bt:Image size="80" resid="blue-icon-80" />
        </Icon>
        <Control xsi:type="Button" id="Contoso.Button2">
            <!-- information on the control omitted -->
        </Control>
        <OfficeControl id="Superscript" />
        <!-- other controls, as needed -->
    </Group>
    <Label resid="customTabLabel1" />
  </CustomTab>
</ExtensionPoint>
```
