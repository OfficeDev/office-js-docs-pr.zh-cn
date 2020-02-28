---
title: 清单文件中的 Action 元素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b05da08f4995c7d8f7270e7fba6f416c9903b066
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324889"
---
# <a name="action-element"></a><span data-ttu-id="89a67-102">Action 元素</span><span class="sxs-lookup"><span data-stu-id="89a67-102">Action element</span></span>

<span data-ttu-id="89a67-103">指定用户选择 [按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件时将执行的操作。</span><span class="sxs-lookup"><span data-stu-id="89a67-103">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="89a67-104">属性</span><span class="sxs-lookup"><span data-stu-id="89a67-104">Attributes</span></span>

|  <span data-ttu-id="89a67-105">属性</span><span class="sxs-lookup"><span data-stu-id="89a67-105">Attribute</span></span>  |  <span data-ttu-id="89a67-106">必需</span><span class="sxs-lookup"><span data-stu-id="89a67-106">Required</span></span>  |  <span data-ttu-id="89a67-107">说明</span><span class="sxs-lookup"><span data-stu-id="89a67-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="89a67-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="89a67-108">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="89a67-109">是</span><span class="sxs-lookup"><span data-stu-id="89a67-109">Yes</span></span>  | <span data-ttu-id="89a67-110">要执行的操作类型</span><span class="sxs-lookup"><span data-stu-id="89a67-110">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="89a67-111">子元素</span><span class="sxs-lookup"><span data-stu-id="89a67-111">Child elements</span></span>

|  <span data-ttu-id="89a67-112">元素</span><span class="sxs-lookup"><span data-stu-id="89a67-112">Element</span></span> |  <span data-ttu-id="89a67-113">说明</span><span class="sxs-lookup"><span data-stu-id="89a67-113">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="89a67-114">FunctionName</span><span class="sxs-lookup"><span data-stu-id="89a67-114">FunctionName</span></span>](#functionname) |    <span data-ttu-id="89a67-115">指定要执行的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="89a67-115">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="89a67-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="89a67-116">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="89a67-117">指定该操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="89a67-117">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="89a67-118"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="89a67-118"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="89a67-119">指定任务窗格容器的 ID。</span><span class="sxs-lookup"><span data-stu-id="89a67-119">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="89a67-120"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="89a67-120"> [Title](#title)</span></span> | <span data-ttu-id="89a67-121">指定任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="89a67-121">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="89a67-122"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="89a67-122"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="89a67-123">指定任务窗格支持固定，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="89a67-123">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="89a67-124">xsi:type</span><span class="sxs-lookup"><span data-stu-id="89a67-124">xsi:type</span></span>

<span data-ttu-id="89a67-p101">此属性指定当用户选择按钮时所执行的操作类型。可取值如下：</span><span class="sxs-lookup"><span data-stu-id="89a67-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="89a67-127">FunctionName</span><span class="sxs-lookup"><span data-stu-id="89a67-127">FunctionName</span></span>

<span data-ttu-id="89a67-p102">**xsi:type** 为“ExecuteFunction”时的必需元素。指定要执行的函数的名称。函数包含在 [FunctionFile](functionfile.md) 元素指定的文件中。</span><span class="sxs-lookup"><span data-stu-id="89a67-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="89a67-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="89a67-131">SourceLocation</span></span>

<span data-ttu-id="89a67-132">**Xsi： type**为 "ShowTaskpane" 时必需的元素。</span><span class="sxs-lookup"><span data-stu-id="89a67-132">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="89a67-133">指定此操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="89a67-133">Specifies the source file location for this action.</span></span> <span data-ttu-id="89a67-134">**resid** 属性必须设置为 **Urls** 元素（位于 **Resources** 元素）中 **Url** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="89a67-134">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="89a67-135">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="89a67-135">TaskpaneId</span></span>

<span data-ttu-id="89a67-136"> **xsi: type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="89a67-136">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="89a67-137">指定任务窗格容器的 ID。</span><span class="sxs-lookup"><span data-stu-id="89a67-137">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="89a67-138">具有多个“ShowTaskpane”操作时，如果想要让每个操作使用独立的窗格，则使用不同的 **TaskpaneId**。</span><span class="sxs-lookup"><span data-stu-id="89a67-138">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="89a67-139">若要让不同的操作共享同一个窗格，则使用同一个 **TaskpaneId**。</span><span class="sxs-lookup"><span data-stu-id="89a67-139">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="89a67-140">当用户选择共享同一个 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将被替换为相应的操作“SourceLocation”。</span><span class="sxs-lookup"><span data-stu-id="89a67-140">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="89a67-141">Outlook 不支持此元素。</span><span class="sxs-lookup"><span data-stu-id="89a67-141">This element is not supported in Outlook.</span></span>

<span data-ttu-id="89a67-142">下面的示例展示了两个共享同一个 **TaskpaneId** 的操作。</span><span class="sxs-lookup"><span data-stu-id="89a67-142">The following example shows two actions that share the same **TaskpaneId**.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

<span data-ttu-id="89a67-p105">下面的示例展示了两个使用不同 **TaskpaneId** 的操作。若要查看上下文中的这些示例，请参阅[简单的外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)。</span><span class="sxs-lookup"><span data-stu-id="89a67-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a><span data-ttu-id="89a67-145">标题</span><span class="sxs-lookup"><span data-stu-id="89a67-145">Title</span></span>

<span data-ttu-id="89a67-146"> **xsi: type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="89a67-146">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="89a67-147">指定此操作任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="89a67-147">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="89a67-148">下面的示例演示使用**Title**元素的操作。</span><span class="sxs-lookup"><span data-stu-id="89a67-148">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="89a67-149">请注意，您不会直接向字符串分配**标题**。</span><span class="sxs-lookup"><span data-stu-id="89a67-149">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="89a67-150">而是为其分配一个资源 ID （resid），该 ID 在清单的 "**资源**" 部分中定义。</span><span class="sxs-lookup"><span data-stu-id="89a67-150">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

## <a name="supportspinning"></a><span data-ttu-id="89a67-151">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="89a67-151">SupportsPinning</span></span>

<span data-ttu-id="89a67-152">**xsi:type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="89a67-152">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="89a67-153">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="89a67-153">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="89a67-154">添加此元素时将值设为 `true` 可以支持任务窗格固定。</span><span class="sxs-lookup"><span data-stu-id="89a67-154">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="89a67-155">这样一来，用户可以“固定”任务窗格，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="89a67-155">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="89a67-156">有关详细信息，请参阅[在 Outlook 中实现可固定的任务窗格](../../outlook/pinnable-taskpane.md)。</span><span class="sxs-lookup"><span data-stu-id="89a67-156">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!NOTE]
> <span data-ttu-id="89a67-157">SupportsPinning 目前仅支持 Outlook 2016 或更高版本（内部版本7628.1000 或更高版本）和 Outlook 2016 或 Mac 上的更高版本（内部版本16.13.503 或更高版本）。</span><span class="sxs-lookup"><span data-stu-id="89a67-157">SupportsPinning is currently only supported by Outlook 2016 or later on Windows (build 7628.1000 or later) and Outlook 2016 or later on Mac (build 16.13.503 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
