---
title: 清单文件中的 OfficeMenu 元素
description: OfficeMenu 元素定义要添加到 Office 上下文菜单中的控件的集合。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d181e0c6f489997a149b9713bdc257f4a2baeb16
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641439"
---
# <a name="officemenu-element"></a><span data-ttu-id="6b425-103">OfficeMenu 元素</span><span class="sxs-lookup"><span data-stu-id="6b425-103">OfficeMenu element</span></span>

<span data-ttu-id="6b425-p101">定义要添加到 Office 上下文菜单的控件集合。适用于 Word、Excel、PowerPoint 和 OneNote 外接程序。</span><span class="sxs-lookup"><span data-stu-id="6b425-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="6b425-106">属性</span><span class="sxs-lookup"><span data-stu-id="6b425-106">Attributes</span></span>

| <span data-ttu-id="6b425-107">属性</span><span class="sxs-lookup"><span data-stu-id="6b425-107">Attribute</span></span>            | <span data-ttu-id="6b425-108">必需</span><span class="sxs-lookup"><span data-stu-id="6b425-108">Required</span></span> | <span data-ttu-id="6b425-109">说明</span><span class="sxs-lookup"><span data-stu-id="6b425-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="6b425-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="6b425-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="6b425-111">是</span><span class="sxs-lookup"><span data-stu-id="6b425-111">Yes</span></span>      | <span data-ttu-id="6b425-112">定义的 OfficeMenu 类型。</span><span class="sxs-lookup"><span data-stu-id="6b425-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="6b425-113">子元素</span><span class="sxs-lookup"><span data-stu-id="6b425-113">Child elements</span></span>

|  <span data-ttu-id="6b425-114">元素</span><span class="sxs-lookup"><span data-stu-id="6b425-114">Element</span></span> |  <span data-ttu-id="6b425-115">必需</span><span class="sxs-lookup"><span data-stu-id="6b425-115">Required</span></span>  |  <span data-ttu-id="6b425-116">说明</span><span class="sxs-lookup"><span data-stu-id="6b425-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6b425-117">Control</span><span class="sxs-lookup"><span data-stu-id="6b425-117">Control</span></span>](#control)    | <span data-ttu-id="6b425-118">是</span><span class="sxs-lookup"><span data-stu-id="6b425-118">Yes</span></span> |  <span data-ttu-id="6b425-119">一个或多个 Control 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="6b425-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="6b425-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="6b425-120">xsi:type</span></span>

<span data-ttu-id="6b425-121">指定要在其中添加此 Office 外接程序的 Office 客户端应用程序的内置菜单。</span><span class="sxs-lookup"><span data-stu-id="6b425-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="6b425-p102">`ContextMenuText` -  当用户选定文本，然后打开（右键单击）选定文本上的上下文菜单时显示上下文菜单上的项。适用于 Word、Excel、PowerPoint 和 OneNote。</span><span class="sxs-lookup"><span data-stu-id="6b425-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="6b425-p103">`ContextMenuCell` -  当用户打开（右键单击）电子表格中的某个单元格上的上下文菜单时显示上下文菜单上的项。适用于 Excel。</span><span class="sxs-lookup"><span data-stu-id="6b425-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span>

## <a name="control"></a><span data-ttu-id="6b425-126">Control</span><span class="sxs-lookup"><span data-stu-id="6b425-126">Control</span></span>

<span data-ttu-id="6b425-127">每个 **OfficeMenu** 元素都需要一个或多个 [menu](control.md#menu-dropdown-button-controls) 控件。</span><span class="sxs-lookup"><span data-stu-id="6b425-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="6b425-128">示例</span><span class="sxs-lookup"><span data-stu-id="6b425-128">Example</span></span>

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
