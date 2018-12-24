---
title: 清单文件中的 Control 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: e5d8574e322c21e768fb9f66fe9bbb0c12a34ed4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433934"
---
# <a name="control-element"></a><span data-ttu-id="de2ec-102">Control 元素</span><span class="sxs-lookup"><span data-stu-id="de2ec-102">Control element</span></span>

<span data-ttu-id="de2ec-p101">定义执行操作或启动任务窗格的 JavaScript 函数。**Control** 元素可以是按钮选项，也可以是菜单选项。[Group](group.md) 元素中至少需包括一个 **Control**。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="de2ec-106">属性</span><span class="sxs-lookup"><span data-stu-id="de2ec-106">Attributes</span></span>

|  <span data-ttu-id="de2ec-107">属性</span><span class="sxs-lookup"><span data-stu-id="de2ec-107">Attribute</span></span>  |  <span data-ttu-id="de2ec-108">必需</span><span class="sxs-lookup"><span data-stu-id="de2ec-108">Required</span></span>  |  <span data-ttu-id="de2ec-109">说明</span><span class="sxs-lookup"><span data-stu-id="de2ec-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="de2ec-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="de2ec-110">**xsi:type**</span></span>|<span data-ttu-id="de2ec-111">必需</span><span class="sxs-lookup"><span data-stu-id="de2ec-111">Yes</span></span>|<span data-ttu-id="de2ec-p102">正在定义的控件类型。可以是 `Button`、`Menu` 或 `MobileButton`。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="de2ec-114">**id**</span><span class="sxs-lookup"><span data-stu-id="de2ec-114">**id**</span></span>|<span data-ttu-id="de2ec-115">否</span><span class="sxs-lookup"><span data-stu-id="de2ec-115">No</span></span>|<span data-ttu-id="de2ec-p103">控件元素的 ID。最多可包含 125 个字符。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="de2ec-118">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。</span><span class="sxs-lookup"><span data-stu-id="de2ec-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing VersionOverrides element must have an  attribute value of .</span></span> <span data-ttu-id="de2ec-119">它只适用于 [MobileFormFactor](mobileformfactor.md) 元素内包含的 **Control** 元素。</span><span class="sxs-lookup"><span data-stu-id="de2ec-119">Note: The  value for xsi:type is defined in VersionOverrides schema 1.1. It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="de2ec-120">按钮控件</span><span class="sxs-lookup"><span data-stu-id="de2ec-120">Button control</span></span>

<span data-ttu-id="de2ec-p105">当用户选择某个按钮时，将执行一个操作。它可以执行函数或显示任务窗格。每个按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="de2ec-124">子元素</span><span class="sxs-lookup"><span data-stu-id="de2ec-124">Child elements</span></span>
|  <span data-ttu-id="de2ec-125">元素</span><span class="sxs-lookup"><span data-stu-id="de2ec-125">Element</span></span> |  <span data-ttu-id="de2ec-126">必需</span><span class="sxs-lookup"><span data-stu-id="de2ec-126">Required</span></span>  |  <span data-ttu-id="de2ec-127">说明</span><span class="sxs-lookup"><span data-stu-id="de2ec-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="de2ec-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="de2ec-128">**Label**</span></span>     | <span data-ttu-id="de2ec-129">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-129">Yes</span></span> |  <span data-ttu-id="de2ec-p106">按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="de2ec-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="de2ec-132">**ToolTip**</span></span>  |<span data-ttu-id="de2ec-133">否</span><span class="sxs-lookup"><span data-stu-id="de2ec-133">No</span></span>|<span data-ttu-id="de2ec-p107">按钮的工具提示。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="de2ec-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="de2ec-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="de2ec-138">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-138">Yes</span></span> |  <span data-ttu-id="de2ec-139">按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="de2ec-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="de2ec-140">Icon</span><span class="sxs-lookup"><span data-stu-id="de2ec-140">Icon</span></span>](icon.md)      | <span data-ttu-id="de2ec-141">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-141">Yes</span></span> |  <span data-ttu-id="de2ec-142">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="de2ec-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="de2ec-143">Action</span><span class="sxs-lookup"><span data-stu-id="de2ec-143">Action</span></span>](action.md)    | <span data-ttu-id="de2ec-144">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-144">Yes</span></span> |  <span data-ttu-id="de2ec-145">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="de2ec-145">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="de2ec-146">ExecuteFunction 按钮示例</span><span class="sxs-lookup"><span data-stu-id="de2ec-146">ExecuteFunction button example</span></span>

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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="de2ec-147">ShowTaskpane 按钮示例</span><span class="sxs-lookup"><span data-stu-id="de2ec-147">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="de2ec-148">菜单（下拉）控件</span><span class="sxs-lookup"><span data-stu-id="de2ec-148">Menu (dropdown button) controls</span></span>

<span data-ttu-id="de2ec-p108">菜单定义选项的静态列表。每个菜单项将执行函数或显示任务窗格。不支持子菜单。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="de2ec-152">使用 **PrimaryCommandSurface** 或 **ContextMenu** [扩展点](extensionpoint.md)时，菜单控件定义：</span><span class="sxs-lookup"><span data-stu-id="de2ec-152">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="de2ec-153">根级别菜单项。</span><span class="sxs-lookup"><span data-stu-id="de2ec-153">A root-level menu item.</span></span>

- <span data-ttu-id="de2ec-154">子菜单项的列表。</span><span class="sxs-lookup"><span data-stu-id="de2ec-154">A list of submenu items.</span></span>

<span data-ttu-id="de2ec-p109">当与  **PrimaryCommandSurface** 一起使用时，根菜单项将显示为功能区上的按钮。选择该按钮后，子菜单将显示为下拉列表。与 **ContextMenu** 一起使用时，具有子菜单的菜单项将被插入到上下文菜单上。在这两种情况下，单个子菜单项可以执行 JavaScript 函数，也可显示任务窗格。这一次仅支持子菜单的一个级别。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="de2ec-p110">下面的示例演示如何定义具有两个子菜单项的菜单项。第一个子菜单项显示任务窗格，而第二个子菜单项运行 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="de2ec-162">子元素</span><span class="sxs-lookup"><span data-stu-id="de2ec-162">Child elements</span></span>

|  <span data-ttu-id="de2ec-163">元素</span><span class="sxs-lookup"><span data-stu-id="de2ec-163">Element</span></span> |  <span data-ttu-id="de2ec-164">必需</span><span class="sxs-lookup"><span data-stu-id="de2ec-164">Required</span></span>  |  <span data-ttu-id="de2ec-165">说明</span><span class="sxs-lookup"><span data-stu-id="de2ec-165">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="de2ec-166">**Label**</span><span class="sxs-lookup"><span data-stu-id="de2ec-166">**Label**</span></span>     | <span data-ttu-id="de2ec-167">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-167">Yes</span></span> |  <span data-ttu-id="de2ec-p111">按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="de2ec-170">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="de2ec-170">**ToolTip**</span></span>  |<span data-ttu-id="de2ec-171">否</span><span class="sxs-lookup"><span data-stu-id="de2ec-171">No</span></span>|<span data-ttu-id="de2ec-p112">按钮的工具提示。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 [Resources](resources.md) 元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="de2ec-175">Supertip</span><span class="sxs-lookup"><span data-stu-id="de2ec-175">Supertip</span></span>](supertip.md)  | <span data-ttu-id="de2ec-176">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-176">Yes</span></span> |  <span data-ttu-id="de2ec-177">此按钮的 supertip。</span><span class="sxs-lookup"><span data-stu-id="de2ec-177">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="de2ec-178">Icon</span><span class="sxs-lookup"><span data-stu-id="de2ec-178">Icon</span></span>](icon.md)      | <span data-ttu-id="de2ec-179">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-179">Yes</span></span> |  <span data-ttu-id="de2ec-180">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="de2ec-180">An image for the button.</span></span>         |
|  <span data-ttu-id="de2ec-181">**Items**</span><span class="sxs-lookup"><span data-stu-id="de2ec-181">**Items**</span></span>     | <span data-ttu-id="de2ec-182">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-182">Yes</span></span> |  <span data-ttu-id="de2ec-p113">菜单中显示的按钮的集合。包含每个子菜单项的 **Item** 元素。每个 **Item** 元素均包含 [按钮控件](#button-control) 的子元素。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="de2ec-186">菜单控件示例</span><span class="sxs-lookup"><span data-stu-id="de2ec-186">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="de2ec-187">MobileButton 控件</span><span class="sxs-lookup"><span data-stu-id="de2ec-187">MobileButton control</span></span>

<span data-ttu-id="de2ec-p114">当用户选择某个移动按钮时，将执行一个操作。它可以执行函数或显示任务窗格。 每个移动按钮控件必须具有对清单唯一的 `id`。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="de2ec-p115">在 VersionOverrides 架构 1.1 中定义了 **xsi:type** 的 `MobileButton` 值。包含  [VersionOverrides](versionoverrides.md) 元素的 `VersionOverridesV1_1` 属性值必须为 `xsi:type`。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="de2ec-193">子元素</span><span class="sxs-lookup"><span data-stu-id="de2ec-193">Child elements</span></span>
|  <span data-ttu-id="de2ec-194">元素</span><span class="sxs-lookup"><span data-stu-id="de2ec-194">Element</span></span> |  <span data-ttu-id="de2ec-195">必需</span><span class="sxs-lookup"><span data-stu-id="de2ec-195">Required</span></span>  |  <span data-ttu-id="de2ec-196">说明</span><span class="sxs-lookup"><span data-stu-id="de2ec-196">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="de2ec-197">**Label**</span><span class="sxs-lookup"><span data-stu-id="de2ec-197">**Label**</span></span>     | <span data-ttu-id="de2ec-198">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-198">Yes</span></span> |  <span data-ttu-id="de2ec-p116">按钮文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 [Resources](resources.md) 元素）中 **String** 元素的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="de2ec-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="de2ec-201">Icon</span><span class="sxs-lookup"><span data-stu-id="de2ec-201">Icon</span></span>](icon.md)      | <span data-ttu-id="de2ec-202">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-202">Yes</span></span> |  <span data-ttu-id="de2ec-203">按钮的图像。</span><span class="sxs-lookup"><span data-stu-id="de2ec-203">An image for the button.</span></span>         |
|  [<span data-ttu-id="de2ec-204">Action</span><span class="sxs-lookup"><span data-stu-id="de2ec-204">Action</span></span>](action.md)    | <span data-ttu-id="de2ec-205">是</span><span class="sxs-lookup"><span data-stu-id="de2ec-205">Yes</span></span> |  <span data-ttu-id="de2ec-206">指定要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="de2ec-206">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="de2ec-207">ExecuteFunction 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="de2ec-207">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="de2ec-208">ShowTaskpane 移动按钮示例</span><span class="sxs-lookup"><span data-stu-id="de2ec-208">ShowTaskpane mobile button example</span></span>

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