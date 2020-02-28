---
title: 清单文件中的 Control 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: ed76cc46c624d1b97d43e4270944b8ef4dc63723
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42323796"
---
# <a name="control-element"></a><span data-ttu-id="a05e7-102">Control 元素</span><span class="sxs-lookup"><span data-stu-id="a05e7-102">Control element</span></span>

<span data-ttu-id="a05e7-p101">定义执行操作或启动任务窗格的 JavaScript 函数。**Control** 元素可以是按钮选项，也可以是菜单选项。[Group](group.md) 元素中至少需包括一个 **Control**。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="a05e7-106">属性</span><span class="sxs-lookup"><span data-stu-id="a05e7-106">Attributes</span></span>

|  <span data-ttu-id="a05e7-107">属性</span><span class="sxs-lookup"><span data-stu-id="a05e7-107">Attribute</span></span>  |  <span data-ttu-id="a05e7-108">必需</span><span class="sxs-lookup"><span data-stu-id="a05e7-108">Required</span></span>  |  <span data-ttu-id="a05e7-109">说明</span><span class="sxs-lookup"><span data-stu-id="a05e7-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="a05e7-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="a05e7-110">**xsi:type**</span></span>|<span data-ttu-id="a05e7-111">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-111">Yes</span></span>|<span data-ttu-id="a05e7-p102">正在定义的控件类型。可以是 `Button`、`Menu` 或 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="a05e7-114">**id**</span><span class="sxs-lookup"><span data-stu-id="a05e7-114">**id**</span></span>|<span data-ttu-id="a05e7-115">否</span><span class="sxs-lookup"><span data-stu-id="a05e7-115">No</span></span>|<span data-ttu-id="a05e7-p103">控件元素的 ID。最多可包含 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="a05e7-118">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。</span><span class="sxs-lookup"><span data-stu-id="a05e7-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="a05e7-119">它只适用于 [MobileFormFactor](mobileformfactor.md) 元素内包含的 **Control** 元素。</span><span class="sxs-lookup"><span data-stu-id="a05e7-119">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="a05e7-120">按钮控件</span><span class="sxs-lookup"><span data-stu-id="a05e7-120">Button control</span></span>

<span data-ttu-id="a05e7-p105">当用户选择某个按钮时，将执行一个操作。它可以执行函数或显示任务窗格。每个按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="a05e7-124">子元素</span><span class="sxs-lookup"><span data-stu-id="a05e7-124">Child elements</span></span>
|  <span data-ttu-id="a05e7-125">元素</span><span class="sxs-lookup"><span data-stu-id="a05e7-125">Element</span></span> |  <span data-ttu-id="a05e7-126">必需</span><span class="sxs-lookup"><span data-stu-id="a05e7-126">Required</span></span>  |  <span data-ttu-id="a05e7-127">说明</span><span class="sxs-lookup"><span data-stu-id="a05e7-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a05e7-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="a05e7-128">**Label**</span></span>     | <span data-ttu-id="a05e7-129">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-129">Yes</span></span> |  <span data-ttu-id="a05e7-130">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="a05e7-130">The text for the button.</span></span> <span data-ttu-id="a05e7-131">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="a05e7-131">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="a05e7-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="a05e7-132">**ToolTip**</span></span>  |<span data-ttu-id="a05e7-133">否</span><span class="sxs-lookup"><span data-stu-id="a05e7-133">No</span></span>|<span data-ttu-id="a05e7-134">按钮的工具提示。</span><span class="sxs-lookup"><span data-stu-id="a05e7-134">The tooltip for the button.</span></span> <span data-ttu-id="a05e7-135">必须将“resid”属性设置为 String 元素的 id 属性值。</span><span class="sxs-lookup"><span data-stu-id="a05e7-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="a05e7-136">**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="a05e7-136">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="a05e7-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="a05e7-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="a05e7-138">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-138">Yes</span></span> |  <span data-ttu-id="a05e7-139">按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="a05e7-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="a05e7-140">图标</span><span class="sxs-lookup"><span data-stu-id="a05e7-140">Icon</span></span>](icon.md)      | <span data-ttu-id="a05e7-141">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-141">Yes</span></span> |  <span data-ttu-id="a05e7-142">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="a05e7-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="a05e7-143">Action</span><span class="sxs-lookup"><span data-stu-id="a05e7-143">Action</span></span>](action.md)    | <span data-ttu-id="a05e7-144">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-144">Yes</span></span> |  <span data-ttu-id="a05e7-145">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="a05e7-145">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="a05e7-146">ExecuteFunction 按钮示例</span><span class="sxs-lookup"><span data-stu-id="a05e7-146">ExecuteFunction button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
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
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="a05e7-147">ShowTaskpane 按钮示例</span><span class="sxs-lookup"><span data-stu-id="a05e7-147">ShowTaskpane button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="a05e7-148">菜单（下拉）控件</span><span class="sxs-lookup"><span data-stu-id="a05e7-148">Menu (dropdown button) controls</span></span>

<span data-ttu-id="a05e7-p108">菜单定义选项的静态列表。每个菜单项将执行函数或显示任务窗格。不支持子菜单。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="a05e7-152">使用 **PrimaryCommandSurface** 或 **ContextMenu** [扩展点](extensionpoint.md)时，菜单控件定义：</span><span class="sxs-lookup"><span data-stu-id="a05e7-152">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="a05e7-153">根级别菜单项。</span><span class="sxs-lookup"><span data-stu-id="a05e7-153">A root-level menu item.</span></span>

- <span data-ttu-id="a05e7-154">子菜单项的列表。</span><span class="sxs-lookup"><span data-stu-id="a05e7-154">A list of submenu items.</span></span>

<span data-ttu-id="a05e7-p109">与 **PrimaryCommandSurface** 结合使用时，根菜单项显示为功能区上的一个按钮。选择此按钮时，子菜单显示为下拉列表。与 **ContextMenu** 结合使用时，将在上下文菜单上插入包含子菜单的菜单项。在这两种情况中，单个子菜单项均可以执行 JavaScript 函数或显示任务窗格。目前只支持一种子菜单级别。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="a05e7-p110">下面的示例演示如何定义具有两个子菜单项的菜单项。第一个子菜单项显示任务窗格，而第二个子菜单项运行 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
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
    <Item id="showGallery2">
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
    <Item id="showGallery3">
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
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

### <a name="child-elements"></a><span data-ttu-id="a05e7-162">子元素</span><span class="sxs-lookup"><span data-stu-id="a05e7-162">Child elements</span></span>

|  <span data-ttu-id="a05e7-163">元素</span><span class="sxs-lookup"><span data-stu-id="a05e7-163">Element</span></span> |  <span data-ttu-id="a05e7-164">必需</span><span class="sxs-lookup"><span data-stu-id="a05e7-164">Required</span></span>  |  <span data-ttu-id="a05e7-165">说明</span><span class="sxs-lookup"><span data-stu-id="a05e7-165">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a05e7-166">**Label**</span><span class="sxs-lookup"><span data-stu-id="a05e7-166">**Label**</span></span>     | <span data-ttu-id="a05e7-167">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-167">Yes</span></span> |  <span data-ttu-id="a05e7-168">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="a05e7-168">The text for the button.</span></span> <span data-ttu-id="a05e7-169">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="a05e7-169">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="a05e7-170">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="a05e7-170">**ToolTip**</span></span>  |<span data-ttu-id="a05e7-171">否</span><span class="sxs-lookup"><span data-stu-id="a05e7-171">No</span></span>|<span data-ttu-id="a05e7-172">按钮的工具提示。</span><span class="sxs-lookup"><span data-stu-id="a05e7-172">The tooltip for the button.</span></span> <span data-ttu-id="a05e7-173">必须将“resid”属性设置为 String 元素的 id 属性值。</span><span class="sxs-lookup"><span data-stu-id="a05e7-173">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="a05e7-174">**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="a05e7-174">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="a05e7-175">Supertip</span><span class="sxs-lookup"><span data-stu-id="a05e7-175">Supertip</span></span>](supertip.md)  | <span data-ttu-id="a05e7-176">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-176">Yes</span></span> |  <span data-ttu-id="a05e7-177">此按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="a05e7-177">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="a05e7-178">Icon</span><span class="sxs-lookup"><span data-stu-id="a05e7-178">Icon</span></span>](icon.md)      | <span data-ttu-id="a05e7-179">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-179">Yes</span></span> |  <span data-ttu-id="a05e7-180">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="a05e7-180">An image for the button.</span></span>         |
|  <span data-ttu-id="a05e7-181">**Items**</span><span class="sxs-lookup"><span data-stu-id="a05e7-181">**Items**</span></span>     | <span data-ttu-id="a05e7-182">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-182">Yes</span></span> |  <span data-ttu-id="a05e7-183">菜单中显示的按钮的集合。</span><span class="sxs-lookup"><span data-stu-id="a05e7-183">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="a05e7-184">包含每个子菜单项的 **Item** 元素。</span><span class="sxs-lookup"><span data-stu-id="a05e7-184">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="a05e7-185">每个 **Item** 元素均包含 [按钮控件](#button-control)的子元素。</span><span class="sxs-lookup"><span data-stu-id="a05e7-185">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="a05e7-186">菜单控件示例</span><span class="sxs-lookup"><span data-stu-id="a05e7-186">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
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
    <Item id="showGallery2">
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
    <Item id="showGallery3">
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
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>

```

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
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
    <Item id="msgReadMenuItem1">
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

## <a name="mobilebutton-control"></a><span data-ttu-id="a05e7-187">MobileButton 控件</span><span class="sxs-lookup"><span data-stu-id="a05e7-187">MobileButton control</span></span>

<span data-ttu-id="a05e7-p114">当用户选择某个移动按钮时，将执行一个操作。它可以执行函数或显示任务窗格。 每个移动按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="a05e7-p115">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="a05e7-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="a05e7-193">子元素</span><span class="sxs-lookup"><span data-stu-id="a05e7-193">Child elements</span></span>
|  <span data-ttu-id="a05e7-194">元素</span><span class="sxs-lookup"><span data-stu-id="a05e7-194">Element</span></span> |  <span data-ttu-id="a05e7-195">必需</span><span class="sxs-lookup"><span data-stu-id="a05e7-195">Required</span></span>  |  <span data-ttu-id="a05e7-196">说明</span><span class="sxs-lookup"><span data-stu-id="a05e7-196">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a05e7-197">**Label**</span><span class="sxs-lookup"><span data-stu-id="a05e7-197">**Label**</span></span>     | <span data-ttu-id="a05e7-198">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-198">Yes</span></span> |  <span data-ttu-id="a05e7-199">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="a05e7-199">The text for the button.</span></span> <span data-ttu-id="a05e7-200">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="a05e7-200">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="a05e7-201">Icon</span><span class="sxs-lookup"><span data-stu-id="a05e7-201">Icon</span></span>](icon.md)      | <span data-ttu-id="a05e7-202">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-202">Yes</span></span> |  <span data-ttu-id="a05e7-203">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="a05e7-203">An image for the button.</span></span>         |
|  [<span data-ttu-id="a05e7-204">Action</span><span class="sxs-lookup"><span data-stu-id="a05e7-204">Action</span></span>](action.md)    | <span data-ttu-id="a05e7-205">是</span><span class="sxs-lookup"><span data-stu-id="a05e7-205">Yes</span></span> |  <span data-ttu-id="a05e7-206">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="a05e7-206">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="a05e7-207">ExecuteFunction 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="a05e7-207">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="a05e7-208">ShowTaskpane 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="a05e7-208">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
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
