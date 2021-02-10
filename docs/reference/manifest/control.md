---
title: 清单文件中的 Control 元素
description: 定义执行操作或启动任务窗格的 JavaScript 函数。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 737902bef52edeb70e2c5760df5bb589b624271b
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173981"
---
# <a name="control-element"></a><span data-ttu-id="ec70f-103">Control 元素</span><span class="sxs-lookup"><span data-stu-id="ec70f-103">Control element</span></span>

<span data-ttu-id="ec70f-p101">定义执行操作或启动任务窗格的 JavaScript 函数。**Control** 元素可以是按钮选项，也可以是菜单选项。[Group](group.md) 元素中至少需包括一个 **Control**。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="ec70f-107">属性</span><span class="sxs-lookup"><span data-stu-id="ec70f-107">Attributes</span></span>

|  <span data-ttu-id="ec70f-108">属性</span><span class="sxs-lookup"><span data-stu-id="ec70f-108">Attribute</span></span>  |  <span data-ttu-id="ec70f-109">必需</span><span class="sxs-lookup"><span data-stu-id="ec70f-109">Required</span></span>  |  <span data-ttu-id="ec70f-110">说明</span><span class="sxs-lookup"><span data-stu-id="ec70f-110">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="ec70f-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="ec70f-111">**xsi:type**</span></span>|<span data-ttu-id="ec70f-112">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-112">Yes</span></span>|<span data-ttu-id="ec70f-p102">正在定义的控件类型。可以是 `Button`、`Menu` 或 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="ec70f-115">**id**</span><span class="sxs-lookup"><span data-stu-id="ec70f-115">**id**</span></span>|<span data-ttu-id="ec70f-116">否</span><span class="sxs-lookup"><span data-stu-id="ec70f-116">No</span></span>|<span data-ttu-id="ec70f-p103">控件元素的 ID。最多可包含 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="ec70f-119">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。</span><span class="sxs-lookup"><span data-stu-id="ec70f-119">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="ec70f-120">它只适用于 [MobileFormFactor](mobileformfactor.md) 元素内包含的 **Control** 元素。</span><span class="sxs-lookup"><span data-stu-id="ec70f-120">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="ec70f-121">按钮控件</span><span class="sxs-lookup"><span data-stu-id="ec70f-121">Button control</span></span>

<span data-ttu-id="ec70f-p105">当用户选择某个按钮时，将执行一个操作。它可以执行函数或显示任务窗格。每个按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="ec70f-125">子元素</span><span class="sxs-lookup"><span data-stu-id="ec70f-125">Child elements</span></span>
|  <span data-ttu-id="ec70f-126">元素</span><span class="sxs-lookup"><span data-stu-id="ec70f-126">Element</span></span> |  <span data-ttu-id="ec70f-127">必需</span><span class="sxs-lookup"><span data-stu-id="ec70f-127">Required</span></span>  |  <span data-ttu-id="ec70f-128">说明</span><span class="sxs-lookup"><span data-stu-id="ec70f-128">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ec70f-129">**Label**</span><span class="sxs-lookup"><span data-stu-id="ec70f-129">**Label**</span></span>     | <span data-ttu-id="ec70f-130">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-130">Yes</span></span> |  <span data-ttu-id="ec70f-131">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="ec70f-131">The text for the button.</span></span> <span data-ttu-id="ec70f-132">**resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)</span><span class="sxs-lookup"><span data-stu-id="ec70f-132">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="ec70f-133">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="ec70f-133">**ToolTip**</span></span>    |<span data-ttu-id="ec70f-134">否</span><span class="sxs-lookup"><span data-stu-id="ec70f-134">No</span></span>|<span data-ttu-id="ec70f-135">按钮的工具提示。</span><span class="sxs-lookup"><span data-stu-id="ec70f-135">The tooltip for the button.</span></span> <span data-ttu-id="ec70f-136">**resid** 属性的长度不能超过 32 个字符，必须设置为 **String** 元素的 **id** 属性值。</span><span class="sxs-lookup"><span data-stu-id="ec70f-136">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="ec70f-137">**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="ec70f-137">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="ec70f-138">Supertip</span><span class="sxs-lookup"><span data-stu-id="ec70f-138">Supertip</span></span>](supertip.md)  | <span data-ttu-id="ec70f-139">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-139">Yes</span></span> |  <span data-ttu-id="ec70f-140">按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="ec70f-140">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="ec70f-141">Icon</span><span class="sxs-lookup"><span data-stu-id="ec70f-141">Icon</span></span>](icon.md)      | <span data-ttu-id="ec70f-142">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-142">Yes</span></span> |  <span data-ttu-id="ec70f-143">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="ec70f-143">An image for the button.</span></span>         |
|  [<span data-ttu-id="ec70f-144">Action</span><span class="sxs-lookup"><span data-stu-id="ec70f-144">Action</span></span>](action.md)    | <span data-ttu-id="ec70f-145">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-145">Yes</span></span> |  <span data-ttu-id="ec70f-146">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="ec70f-146">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="ec70f-147">Enabled</span><span class="sxs-lookup"><span data-stu-id="ec70f-147">Enabled</span></span>](enabled.md)    | <span data-ttu-id="ec70f-148">否</span><span class="sxs-lookup"><span data-stu-id="ec70f-148">No</span></span> |  <span data-ttu-id="ec70f-149">指定加载项启动时是否启用控件。</span><span class="sxs-lookup"><span data-stu-id="ec70f-149">Specifies whether the control is enabled when the add-in launches.</span></span>  |
|  [<span data-ttu-id="ec70f-150">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="ec70f-150">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="ec70f-151">否</span><span class="sxs-lookup"><span data-stu-id="ec70f-151">No</span></span> |  <span data-ttu-id="ec70f-152">指定该按钮是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。</span><span class="sxs-lookup"><span data-stu-id="ec70f-152">Specifies whether the button should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="ec70f-153">如果使用，则它必须是第 *一个* 子元素。</span><span class="sxs-lookup"><span data-stu-id="ec70f-153">If used, it must be the *first* child element.</span></span> |

### <a name="executefunction-button-example"></a><span data-ttu-id="ec70f-154">ExecuteFunction 按钮示例</span><span class="sxs-lookup"><span data-stu-id="ec70f-154">ExecuteFunction button example</span></span>

<span data-ttu-id="ec70f-155">在下面的示例中，当加载项启动时禁用该按钮。</span><span class="sxs-lookup"><span data-stu-id="ec70f-155">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="ec70f-156">可以编程方式启用它。</span><span class="sxs-lookup"><span data-stu-id="ec70f-156">It can be programmatically enabled.</span></span> <span data-ttu-id="ec70f-157">有关详细信息，请参阅[启用和禁用加载项命令](../../design/disable-add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="ec70f-157">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="ec70f-158">ShowTaskpane 按钮示例</span><span class="sxs-lookup"><span data-stu-id="ec70f-158">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="ec70f-159">菜单（下拉）控件</span><span class="sxs-lookup"><span data-stu-id="ec70f-159">Menu (dropdown button) controls</span></span>

<span data-ttu-id="ec70f-p110">菜单定义选项的静态列表。每个菜单项将执行函数或显示任务窗格。不支持子菜单。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p110">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="ec70f-163">使用 **PrimaryCommandSurface** 或 **ContextMenu** [扩展点](extensionpoint.md)时，菜单控件定义：</span><span class="sxs-lookup"><span data-stu-id="ec70f-163">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="ec70f-164">根级别菜单项。</span><span class="sxs-lookup"><span data-stu-id="ec70f-164">A root-level menu item.</span></span>

- <span data-ttu-id="ec70f-165">子菜单项的列表。</span><span class="sxs-lookup"><span data-stu-id="ec70f-165">A list of submenu items.</span></span>

<span data-ttu-id="ec70f-p111">与 **PrimaryCommandSurface** 结合使用时，根菜单项显示为功能区上的一个按钮。选择此按钮时，子菜单显示为下拉列表。与 **ContextMenu** 结合使用时，将在上下文菜单上插入包含子菜单的菜单项。在这两种情况中，单个子菜单项均可以执行 JavaScript 函数或显示任务窗格。目前只支持一种子菜单级别。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p111">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="ec70f-p112">下面的示例演示如何定义具有两个子菜单项的菜单项。第一个子菜单项显示任务窗格，而第二个子菜单项运行 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p112">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="ec70f-173">子元素</span><span class="sxs-lookup"><span data-stu-id="ec70f-173">Child elements</span></span>

|  <span data-ttu-id="ec70f-174">元素</span><span class="sxs-lookup"><span data-stu-id="ec70f-174">Element</span></span> |  <span data-ttu-id="ec70f-175">必需</span><span class="sxs-lookup"><span data-stu-id="ec70f-175">Required</span></span>  |  <span data-ttu-id="ec70f-176">说明</span><span class="sxs-lookup"><span data-stu-id="ec70f-176">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ec70f-177">**Label**</span><span class="sxs-lookup"><span data-stu-id="ec70f-177">**Label**</span></span>     | <span data-ttu-id="ec70f-178">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-178">Yes</span></span> |  <span data-ttu-id="ec70f-179">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="ec70f-179">The text for the button.</span></span> <span data-ttu-id="ec70f-180">**resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)</span><span class="sxs-lookup"><span data-stu-id="ec70f-180">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="ec70f-181">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="ec70f-181">**ToolTip**</span></span>    |<span data-ttu-id="ec70f-182">否</span><span class="sxs-lookup"><span data-stu-id="ec70f-182">No</span></span>|<span data-ttu-id="ec70f-183">按钮的工具提示。</span><span class="sxs-lookup"><span data-stu-id="ec70f-183">The tooltip for the button.</span></span> <span data-ttu-id="ec70f-184">**resid** 属性的长度不能超过 32 个字符，必须设置为 **String** 元素的 **id** 属性值。</span><span class="sxs-lookup"><span data-stu-id="ec70f-184">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="ec70f-185">**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="ec70f-185">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="ec70f-186">Supertip</span><span class="sxs-lookup"><span data-stu-id="ec70f-186">Supertip</span></span>](supertip.md)  | <span data-ttu-id="ec70f-187">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-187">Yes</span></span> |  <span data-ttu-id="ec70f-188">此按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="ec70f-188">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="ec70f-189">Icon</span><span class="sxs-lookup"><span data-stu-id="ec70f-189">Icon</span></span>](icon.md)      | <span data-ttu-id="ec70f-190">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-190">Yes</span></span> |  <span data-ttu-id="ec70f-191">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="ec70f-191">An image for the button.</span></span>         |
|  <span data-ttu-id="ec70f-192">**Items**</span><span class="sxs-lookup"><span data-stu-id="ec70f-192">**Items**</span></span>     | <span data-ttu-id="ec70f-193">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-193">Yes</span></span> |  <span data-ttu-id="ec70f-194">菜单中显示的按钮的集合。</span><span class="sxs-lookup"><span data-stu-id="ec70f-194">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="ec70f-195">包含每个子菜单项的 **Item** 元素。</span><span class="sxs-lookup"><span data-stu-id="ec70f-195">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="ec70f-196">每个 **Item** 元素均包含 [按钮控件](#button-control) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="ec70f-196">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|
|  [<span data-ttu-id="ec70f-197">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="ec70f-197">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="ec70f-198">否</span><span class="sxs-lookup"><span data-stu-id="ec70f-198">No</span></span> |  <span data-ttu-id="ec70f-199">指定菜单是否应该显示在支持自定义上下文选项卡的应用程序和平台组合上。</span><span class="sxs-lookup"><span data-stu-id="ec70f-199">Specifies whether the menu should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="ec70f-200">如果使用，则它必须是第 *一个* 子元素。</span><span class="sxs-lookup"><span data-stu-id="ec70f-200">If used, it must be the *first* child element.</span></span> |

### <a name="menu-control-examples"></a><span data-ttu-id="ec70f-201">菜单控件示例</span><span class="sxs-lookup"><span data-stu-id="ec70f-201">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
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

## <a name="mobilebutton-control"></a><span data-ttu-id="ec70f-202">MobileButton 控件</span><span class="sxs-lookup"><span data-stu-id="ec70f-202">MobileButton control</span></span>

<span data-ttu-id="ec70f-p117">当用户选择某个移动按钮时，将执行一个操作。它可以执行函数或显示任务窗格。 每个移动按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p117">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="ec70f-p118">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="ec70f-p118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="ec70f-208">子元素</span><span class="sxs-lookup"><span data-stu-id="ec70f-208">Child elements</span></span>
|  <span data-ttu-id="ec70f-209">元素</span><span class="sxs-lookup"><span data-stu-id="ec70f-209">Element</span></span> |  <span data-ttu-id="ec70f-210">必需</span><span class="sxs-lookup"><span data-stu-id="ec70f-210">Required</span></span>  |  <span data-ttu-id="ec70f-211">说明</span><span class="sxs-lookup"><span data-stu-id="ec70f-211">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ec70f-212">**Label**</span><span class="sxs-lookup"><span data-stu-id="ec70f-212">**Label**</span></span>     | <span data-ttu-id="ec70f-213">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-213">Yes</span></span> |  <span data-ttu-id="ec70f-214">按钮文本。</span><span class="sxs-lookup"><span data-stu-id="ec70f-214">The text for the button.</span></span> <span data-ttu-id="ec70f-215">**resid** 属性的长度不能超过 32 个字符，必须设置为 Resources 元素 **中 ShortStrings** 元素 **中 String** 元素的 **id** [属性值。](resources.md)</span><span class="sxs-lookup"><span data-stu-id="ec70f-215">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="ec70f-216">Icon</span><span class="sxs-lookup"><span data-stu-id="ec70f-216">Icon</span></span>](icon.md)      | <span data-ttu-id="ec70f-217">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-217">Yes</span></span> |  <span data-ttu-id="ec70f-218">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="ec70f-218">An image for the button.</span></span>         |
|  [<span data-ttu-id="ec70f-219">Action</span><span class="sxs-lookup"><span data-stu-id="ec70f-219">Action</span></span>](action.md)    | <span data-ttu-id="ec70f-220">是</span><span class="sxs-lookup"><span data-stu-id="ec70f-220">Yes</span></span> |  <span data-ttu-id="ec70f-221">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="ec70f-221">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="ec70f-222">ExecuteFunction 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="ec70f-222">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="ec70f-223">ShowTaskpane 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="ec70f-223">ShowTaskpane mobile button example</span></span>

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
