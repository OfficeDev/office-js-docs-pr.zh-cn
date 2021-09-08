---
title: 清单文件中的 OfficeMenu 元素
description: OfficeMenu 元素定义要添加到上下文菜单的控件Office集合。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d181e0c6f489997a149b9713bdc257f4a2baeb16
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938995"
---
# <a name="officemenu-element"></a>OfficeMenu 元素

定义要添加到 Office 上下文菜单的控件集合。适用于 Word、Excel、PowerPoint 和 OneNote 外接程序。

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

每个 **OfficeMenu** 元素都需要一个或多个 [menu](control.md#menu-dropdown-button-controls) 控件。 

## <a name="example"></a>示例

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
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
