---
title: 清单文件中 MobileButton 类型的 Control 元素
description: 定义移动设备上用于执行操作或启动任务窗格的按钮。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: d498b728bf7f19cf239ffc6178f19cdf9a62de58
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467903"
---
# <a name="control-element-of-type-mobilebutton"></a>类型为 MobileButton 的 Control 元素

定义一个按钮，该按钮执行操作或启动任务窗格，并且该按钮仅出现在移动平台上。

> [!NOTE]
> 本文假定熟悉基本的 [Control 参考文章](control.md) ，该文章包含有关元素属性的重要信息。

当用户选择某个移动按钮时，将执行一个操作。 它可以执行函数或显示任务窗格。 每个按钮控件必须具有在 `id` 清单中所有 **Control** 元素中唯一的属性值。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.1

在 VersionOverrides 架构 1.1 中定义了 `MobileButton` 的  值。包含  [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Label](#label)     | 是 |  按钮文本。 |
|  [Icon](icon.md)      | 是 |  按钮的图像。         |
|  [Action](action.md)    | 是 |  指定要执行的操作。 Control 元素只能有 **一** 个 **Action** 子元素。 |

### <a name="label"></a>标签

通过按钮的唯一属性 **resid** 指定按钮的文本，该属性不能超过 32 个字符，并且必须设置为 [](resources.md) **ShortStrings 元素的 ShortStrings** 子元素中 **String** 元素的 **id** 属性的值。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="examples"></a>示例

在下面的示例中，按钮执行函数。

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

在下面的示例中，该按钮显示任务窗格。

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
