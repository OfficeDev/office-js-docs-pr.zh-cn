---
title: 清单文件中的 OfficeMenu 元素
description: OfficeMenu 元素定义要添加到上下文菜单的控件Office集合。
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 11b68edaef4044fb7ddde0d413debc0339b15c3a
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467741"
---
# <a name="officemenu-element"></a>OfficeMenu 元素

定义要添加到 Office 上下文菜单的控件集合。适用于 Word、Excel、PowerPoint 和 OneNote 外接程序。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

## <a name="attributes"></a>属性

| 属性            | 必需 | 说明                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | 是      | 定义的 OfficeMenu 类型。|

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Control](#control)    | 是 |  一个或多个 Control 对象的集合。  |

## <a name="xsitype"></a>xsi:type

指定要在其中添加此 Office 外接程序的 Office 客户端应用程序的内置菜单。

- `ContextMenuText` -  当用户选定文本，然后打开（右键单击）选定文本上的上下文菜单时显示上下文菜单上的项。适用于 Word、Excel、PowerPoint 和 OneNote。
- `ContextMenuCell` -  当用户打开（右键单击）电子表格中的某个单元格上的上下文菜单时显示上下文菜单上的项。适用于 Excel。

## <a name="control"></a>Control

每个 **OfficeMenu** 元素都需要一个或多个 [Menu 控件](control-menu.md)。 

## <a name="example"></a>示例

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.myMenu">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />
          </Action>
        </Item>
      </Items>
    </Control>
</OfficeMenu>
```
