---
title: 清单文件中 Button 类型的 Control 元素
description: 定义用于执行操作或启动任务窗格的按钮。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: adc58424fe9898bffcbd9e16bed8f3b13b9df4a2
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467883"
---
# <a name="control-element-of-type-button"></a>类型为 Button 的 Control 元素

定义用于执行操作或启动任务窗格的按钮。

> [!NOTE]
> 本文假定熟悉基本的 [Control 参考文章](control.md) ，该文章包含有关元素属性的重要信息。

当用户选择某个按钮时，将执行一个操作。 它可以执行函数或显示任务窗格。 每个按钮控件必须具有在 `id` 清单中所有 **Control** 元素中唯一的属性值。

> [!IMPORTANT]
> 移动平台上忽略"按钮"类型控件。 若要支持移动平台，还必须对"Button"类型的每个控件具有"MobileButton"类型的控件。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Label](#label)     | 是 |  按钮文本。 |
|  **ToolTip**    |否|按钮的工具提示。 **resid** 属性的长度不能超过 32 个字符，必须设置为 **String** 元素 **的 id** 属性的值。 **String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。|
|  [Supertip](supertip.md)  | 是 |  按钮的 supertip。    |
|  [Icon](icon.md)      | 是 |  按钮的图像。         |
|  [Action](action.md)    | 是 |  指定要执行的操作。 Control 元素只能有 **一** 个 **Action** 子元素。 |
|  [Enabled](enabled.md)    | 否 |  指定加载项启动时是否启用控件。  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | 否 |  指定该按钮是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。 如果使用，则它必须是第 *一个子* 元素。 |

### <a name="label"></a>标签

通过按钮的唯一属性 **resid** 指定按钮的文本，该属性不能超过 32 个字符，并且必须设置为 [](resources.md) **ShortStrings 元素的 ShortStrings** 子元素中 **String** 元素的 **id** 属性的值。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- 当父 **VersionOverrides** 的类型为 Taskpane [1.0 时，AddinCommands](../requirement-sets/add-in-commands-requirement-sets.md) 1.1。
- [当父](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) **VersionOverrides** 类型为 Mail 1.0 时，邮箱 1.3。
- [当父](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) **VersionOverrides** 类型为 Mail 1.1 时，邮箱 1.5。

## <a name="examples"></a>示例

在下面的示例中，按钮执行函数。 它还配置为在外接程序启动时禁用。 可以编程方式启用它。 有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。

```xml
<Control xsi:type="Button" id="Contoso.msgReadFunctionButton">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
  <Enabled>false</Enabled>
</Control>
```

在下面的示例中，该按钮显示任务窗格。

```xml
<Control xsi:type="Button" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
