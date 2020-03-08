---
title: 清单文件中的 Control 元素
description: ''
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: ccf7c3065db13a311825498292713b619f1cd745
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/07/2020
ms.locfileid: "42562087"
---
# <a name="control-element"></a><span data-ttu-id="c7f13-102">Control 元素</span><span class="sxs-lookup"><span data-stu-id="c7f13-102">Control element</span></span>

<span data-ttu-id="c7f13-p101">定义执行操作或启动任务窗格的 JavaScript 函数。**Control** 元素可以是按钮选项，也可以是菜单选项。[Group](group.md) 元素中至少需包括一个 **Control**。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="c7f13-106">属性</span><span class="sxs-lookup"><span data-stu-id="c7f13-106">Attributes</span></span>

|  <span data-ttu-id="c7f13-107">属性</span><span class="sxs-lookup"><span data-stu-id="c7f13-107">Attribute</span></span>  |  <span data-ttu-id="c7f13-108">必需</span><span class="sxs-lookup"><span data-stu-id="c7f13-108">Required</span></span>  |  <span data-ttu-id="c7f13-109">说明</span><span class="sxs-lookup"><span data-stu-id="c7f13-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="c7f13-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="c7f13-110">**xsi:type**</span></span>|<span data-ttu-id="c7f13-111">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-111">Yes</span></span>|<span data-ttu-id="c7f13-p102">正在定义的控件类型。可以是 `Button`、`Menu` 或 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="c7f13-114">**id**</span><span class="sxs-lookup"><span data-stu-id="c7f13-114">**id**</span></span>|<span data-ttu-id="c7f13-115">否</span><span class="sxs-lookup"><span data-stu-id="c7f13-115">No</span></span>|<span data-ttu-id="c7f13-p103">控件元素的 ID。最多可包含 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="c7f13-118">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。</span><span class="sxs-lookup"><span data-stu-id="c7f13-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="c7f13-119">它只适用于 [MobileFormFactor](mobileformfactor.md) 元素内包含的 **Control** 元素。</span><span class="sxs-lookup"><span data-stu-id="c7f13-119">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="c7f13-120">按钮控件</span><span class="sxs-lookup"><span data-stu-id="c7f13-120">Button control</span></span>

<span data-ttu-id="c7f13-p105">当用户选择某个按钮时，将执行一个操作。它可以执行函数或显示任务窗格。每个按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="c7f13-124">子元素</span><span class="sxs-lookup"><span data-stu-id="c7f13-124">Child elements</span></span>
|  <span data-ttu-id="c7f13-125">元素</span><span class="sxs-lookup"><span data-stu-id="c7f13-125">Element</span></span> |  <span data-ttu-id="c7f13-126">必需</span><span class="sxs-lookup"><span data-stu-id="c7f13-126">Required</span></span>  |  <span data-ttu-id="c7f13-127">说明</span><span class="sxs-lookup"><span data-stu-id="c7f13-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c7f13-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="c7f13-128">**Label**</span></span>     | <span data-ttu-id="c7f13-129">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-129">Yes</span></span> |  <span data-ttu-id="c7f13-130">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="c7f13-130">The text for the button.</span></span> <span data-ttu-id="c7f13-131">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="c7f13-131">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="c7f13-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="c7f13-132">**ToolTip**</span></span>  |<span data-ttu-id="c7f13-133">否</span><span class="sxs-lookup"><span data-stu-id="c7f13-133">No</span></span>|<span data-ttu-id="c7f13-134">按钮的工具提示。</span><span class="sxs-lookup"><span data-stu-id="c7f13-134">The tooltip for the button.</span></span> <span data-ttu-id="c7f13-135">必须将“resid”属性设置为 String 元素的 id 属性值。</span><span class="sxs-lookup"><span data-stu-id="c7f13-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c7f13-136">**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="c7f13-136">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="c7f13-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="c7f13-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="c7f13-138">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-138">Yes</span></span> |  <span data-ttu-id="c7f13-139">按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="c7f13-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="c7f13-140">图标</span><span class="sxs-lookup"><span data-stu-id="c7f13-140">Icon</span></span>](icon.md)      | <span data-ttu-id="c7f13-141">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-141">Yes</span></span> |  <span data-ttu-id="c7f13-142">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="c7f13-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="c7f13-143">Action</span><span class="sxs-lookup"><span data-stu-id="c7f13-143">Action</span></span>](action.md)    | <span data-ttu-id="c7f13-144">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-144">Yes</span></span> |  <span data-ttu-id="c7f13-145">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="c7f13-145">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="c7f13-146">Enabled</span><span class="sxs-lookup"><span data-stu-id="c7f13-146">Enabled</span></span>](enabled.md)    | <span data-ttu-id="c7f13-147">否</span><span class="sxs-lookup"><span data-stu-id="c7f13-147">No</span></span> |  <span data-ttu-id="c7f13-148">指定在加载项启动时是否启用控件。</span><span class="sxs-lookup"><span data-stu-id="c7f13-148">Specifies whether the control is enabled when the add-in launches.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="c7f13-149">ExecuteFunction 按钮示例</span><span class="sxs-lookup"><span data-stu-id="c7f13-149">ExecuteFunction button example</span></span>

<span data-ttu-id="c7f13-150">在下面的示例中，在加载项启动时，该按钮将被禁用。</span><span class="sxs-lookup"><span data-stu-id="c7f13-150">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="c7f13-151">可以通过编程方式启用它。</span><span class="sxs-lookup"><span data-stu-id="c7f13-151">It can be programmatically enabled.</span></span> <span data-ttu-id="c7f13-152">有关详细信息，请参阅[Enable And Disable 外接程序命令](/office/dev/add-ins/design/disable-add-in-commands)。</span><span class="sxs-lookup"><span data-stu-id="c7f13-152">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

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
  <Enabled>false</Enabled>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="c7f13-153">ShowTaskpane 按钮示例</span><span class="sxs-lookup"><span data-stu-id="c7f13-153">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="c7f13-154">菜单（下拉）控件</span><span class="sxs-lookup"><span data-stu-id="c7f13-154">Menu (dropdown button) controls</span></span>

<span data-ttu-id="c7f13-p109">菜单定义选项的静态列表。每个菜单项将执行函数或显示任务窗格。不支持子菜单。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p109">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="c7f13-158">使用 **PrimaryCommandSurface** 或 **ContextMenu** [扩展点](extensionpoint.md)时，菜单控件定义：</span><span class="sxs-lookup"><span data-stu-id="c7f13-158">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="c7f13-159">根级别菜单项。</span><span class="sxs-lookup"><span data-stu-id="c7f13-159">A root-level menu item.</span></span>

- <span data-ttu-id="c7f13-160">子菜单项的列表。</span><span class="sxs-lookup"><span data-stu-id="c7f13-160">A list of submenu items.</span></span>

<span data-ttu-id="c7f13-p110">与 **PrimaryCommandSurface** 结合使用时，根菜单项显示为功能区上的一个按钮。选择此按钮时，子菜单显示为下拉列表。与 **ContextMenu** 结合使用时，将在上下文菜单上插入包含子菜单的菜单项。在这两种情况中，单个子菜单项均可以执行 JavaScript 函数或显示任务窗格。目前只支持一种子菜单级别。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p110">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="c7f13-p111">下面的示例演示如何定义具有两个子菜单项的菜单项。第一个子菜单项显示任务窗格，而第二个子菜单项运行 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p111">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="c7f13-168">子元素</span><span class="sxs-lookup"><span data-stu-id="c7f13-168">Child elements</span></span>

|  <span data-ttu-id="c7f13-169">元素</span><span class="sxs-lookup"><span data-stu-id="c7f13-169">Element</span></span> |  <span data-ttu-id="c7f13-170">必需</span><span class="sxs-lookup"><span data-stu-id="c7f13-170">Required</span></span>  |  <span data-ttu-id="c7f13-171">说明</span><span class="sxs-lookup"><span data-stu-id="c7f13-171">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c7f13-172">**Label**</span><span class="sxs-lookup"><span data-stu-id="c7f13-172">**Label**</span></span>     | <span data-ttu-id="c7f13-173">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-173">Yes</span></span> |  <span data-ttu-id="c7f13-174">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="c7f13-174">The text for the button.</span></span> <span data-ttu-id="c7f13-175">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="c7f13-175">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="c7f13-176">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="c7f13-176">**ToolTip**</span></span>  |<span data-ttu-id="c7f13-177">否</span><span class="sxs-lookup"><span data-stu-id="c7f13-177">No</span></span>|<span data-ttu-id="c7f13-178">按钮的工具提示。</span><span class="sxs-lookup"><span data-stu-id="c7f13-178">The tooltip for the button.</span></span> <span data-ttu-id="c7f13-179">必须将“resid”属性设置为 String 元素的 id 属性值。</span><span class="sxs-lookup"><span data-stu-id="c7f13-179">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="c7f13-180">**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="c7f13-180">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="c7f13-181">Supertip</span><span class="sxs-lookup"><span data-stu-id="c7f13-181">Supertip</span></span>](supertip.md)  | <span data-ttu-id="c7f13-182">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-182">Yes</span></span> |  <span data-ttu-id="c7f13-183">此按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="c7f13-183">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="c7f13-184">Icon</span><span class="sxs-lookup"><span data-stu-id="c7f13-184">Icon</span></span>](icon.md)      | <span data-ttu-id="c7f13-185">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-185">Yes</span></span> |  <span data-ttu-id="c7f13-186">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="c7f13-186">An image for the button.</span></span>         |
|  <span data-ttu-id="c7f13-187">**Items**</span><span class="sxs-lookup"><span data-stu-id="c7f13-187">**Items**</span></span>     | <span data-ttu-id="c7f13-188">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-188">Yes</span></span> |  <span data-ttu-id="c7f13-189">菜单中显示的按钮的集合。</span><span class="sxs-lookup"><span data-stu-id="c7f13-189">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="c7f13-190">包含每个子菜单项的 **Item** 元素。</span><span class="sxs-lookup"><span data-stu-id="c7f13-190">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="c7f13-191">每个 **Item** 元素均包含 [按钮控件](#button-control)的子元素。</span><span class="sxs-lookup"><span data-stu-id="c7f13-191">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="c7f13-192">菜单控件示例</span><span class="sxs-lookup"><span data-stu-id="c7f13-192">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="c7f13-193">MobileButton 控件</span><span class="sxs-lookup"><span data-stu-id="c7f13-193">MobileButton control</span></span>

<span data-ttu-id="c7f13-p115">当用户选择某个移动按钮时，将执行一个操作。它可以执行函数或显示任务窗格。 每个移动按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p115">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="c7f13-p116">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="c7f13-p116">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="c7f13-199">子元素</span><span class="sxs-lookup"><span data-stu-id="c7f13-199">Child elements</span></span>
|  <span data-ttu-id="c7f13-200">元素</span><span class="sxs-lookup"><span data-stu-id="c7f13-200">Element</span></span> |  <span data-ttu-id="c7f13-201">必需</span><span class="sxs-lookup"><span data-stu-id="c7f13-201">Required</span></span>  |  <span data-ttu-id="c7f13-202">说明</span><span class="sxs-lookup"><span data-stu-id="c7f13-202">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c7f13-203">**Label**</span><span class="sxs-lookup"><span data-stu-id="c7f13-203">**Label**</span></span>     | <span data-ttu-id="c7f13-204">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-204">Yes</span></span> |  <span data-ttu-id="c7f13-205">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="c7f13-205">The text for the button.</span></span> <span data-ttu-id="c7f13-206">**Resid**属性必须设置为[Resources](resources.md)元素中的**ShortStrings**元素中**String**元素的**id**属性的值。</span><span class="sxs-lookup"><span data-stu-id="c7f13-206">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="c7f13-207">Icon</span><span class="sxs-lookup"><span data-stu-id="c7f13-207">Icon</span></span>](icon.md)      | <span data-ttu-id="c7f13-208">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-208">Yes</span></span> |  <span data-ttu-id="c7f13-209">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="c7f13-209">An image for the button.</span></span>         |
|  [<span data-ttu-id="c7f13-210">Action</span><span class="sxs-lookup"><span data-stu-id="c7f13-210">Action</span></span>](action.md)    | <span data-ttu-id="c7f13-211">是</span><span class="sxs-lookup"><span data-stu-id="c7f13-211">Yes</span></span> |  <span data-ttu-id="c7f13-212">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="c7f13-212">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="c7f13-213">ExecuteFunction 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="c7f13-213">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="c7f13-214">ShowTaskpane 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="c7f13-214">ShowTaskpane mobile button example</span></span>

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
