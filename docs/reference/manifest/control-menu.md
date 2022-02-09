---
title: 清单文件中 Menu 类型的 Control 元素
description: 定义其项可以执行操作或启动任务窗格的菜单。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7287b8e2cdf2378140ef50a41306820a0fd4002f
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467895"
---
# <a name="control-element-of-type-menu"></a>Menu 类型的 Control 元素

菜单定义选项列表。 每个菜单项将执行函数或显示任务窗格。

> [!NOTE]
> 本文假定熟悉基本的 [Control 参考文章](control.md) ，该文章包含有关元素属性的重要信息。

菜单控件定义：

- 根级别菜单控件。
- 菜单项的列表。

当与 **PrimaryCommandSurface** 扩展 [点](extensionpoint.md)一同使用时，根菜单项在功能区上显示为按钮。 选择此按钮后，菜单将显示为下拉列表。 不支持子菜单。

当与 **ContextMenu** [扩展点一同使用](extensionpoint.md)时，根菜单项会显示在上下文菜单上。 选择根项目后，菜单项将显示为子菜单。 由于仅支持一个级别的子菜单项，因此任何项自身均不能是子菜单项。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Label](#label)     | 是 |  菜单的文本。 |
|  **ToolTip**    |否|菜单的工具提示。 **resid** 属性的长度不能超过 32 个字符，必须设置为 **String** 元素 **的 id** 属性的值。 **String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。|
|  [Supertip](supertip.md)  | 是 |  此菜单的 supertip。    |
|  [Icon](icon.md)      | 是 |  菜单的图像。         |
|  **Items**     | 是 |  要显示在菜单中的项目的集合。 包含 **每个项目的 Item** 元素。 |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | 否 |  指定菜单是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。 如果使用，则它必须是第 *一个子* 元素。 |

### <a name="label"></a>标签

通过菜单名称的唯一属性 [](resources.md) **resid** 指定文本，该属性不能超过 32 个字符，并且必须设置为 **ShortStrings 元素的 ShortStrings** 子元素中 **String** 元素的 **id** 属性的值。

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

在下面的示例中，该菜单有两个项目。 第一个显示任务窗格。 第二个函数执行函数。 当加载项在支持上下文选项卡的平台上运行时，菜单已配置为不可见。 有关详细信息，请阅读在不支持自定义 [上下文选项卡时实现备用 UI 体验](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

```xml
<Control xsi:type="Menu" id="Contoso.TestMenu2">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="ShowMainTaskPane">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="GetData">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getData</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

在下面的示例中，当加载项在支持上下文选项卡的平台上运行时，菜单的第二项被配置为不可见。 有关详细信息，请阅读在不支持自定义 [上下文选项卡时实现备用 UI 体验](../../design/contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported)。

```xml
<Control xsi:type="Menu" id="Contoso.msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="ShowMainTaskPane">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="msgReadMenuItem1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```
