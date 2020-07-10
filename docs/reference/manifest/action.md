---
title: 清单文件中的 Action 元素
description: 此元素指定当用户选择按钮或菜单控件时要执行的操作。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 92c783a15d104aba0adb722ab887391b4511ebed
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094447"
---
# <a name="action-element"></a><span data-ttu-id="1078f-103">Action 元素</span><span class="sxs-lookup"><span data-stu-id="1078f-103">Action element</span></span>

<span data-ttu-id="1078f-104">指定当用户选择[按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件时要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="1078f-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="1078f-105">属性</span><span class="sxs-lookup"><span data-stu-id="1078f-105">Attributes</span></span>

|  <span data-ttu-id="1078f-106">属性</span><span class="sxs-lookup"><span data-stu-id="1078f-106">Attribute</span></span>  |  <span data-ttu-id="1078f-107">必需</span><span class="sxs-lookup"><span data-stu-id="1078f-107">Required</span></span>  |  <span data-ttu-id="1078f-108">说明</span><span class="sxs-lookup"><span data-stu-id="1078f-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1078f-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1078f-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="1078f-110">是</span><span class="sxs-lookup"><span data-stu-id="1078f-110">Yes</span></span>  | <span data-ttu-id="1078f-111">要执行的操作类型</span><span class="sxs-lookup"><span data-stu-id="1078f-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="1078f-112">子元素</span><span class="sxs-lookup"><span data-stu-id="1078f-112">Child elements</span></span>

|  <span data-ttu-id="1078f-113">元素</span><span class="sxs-lookup"><span data-stu-id="1078f-113">Element</span></span> |  <span data-ttu-id="1078f-114">说明</span><span class="sxs-lookup"><span data-stu-id="1078f-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="1078f-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="1078f-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="1078f-116">指定要执行的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="1078f-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="1078f-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1078f-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="1078f-118">指定该操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="1078f-118">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="1078f-119"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="1078f-119"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="1078f-120">指定任务窗格容器的 ID。</span><span class="sxs-lookup"><span data-stu-id="1078f-120">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="1078f-121"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="1078f-121"> [Title](#title)</span></span> | <span data-ttu-id="1078f-122">指定任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="1078f-122">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="1078f-123"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="1078f-123"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="1078f-124">指定任务窗格支持固定，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="1078f-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="1078f-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="1078f-125">xsi:type</span></span>

<span data-ttu-id="1078f-126">This attribute specifies the kind of action performed when the user selects the button.</span><span class="sxs-lookup"><span data-stu-id="1078f-126">This attribute specifies the kind of action performed when the user selects the button.</span></span> <span data-ttu-id="1078f-127">It can be one of the following:</span><span class="sxs-lookup"><span data-stu-id="1078f-127">It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="1078f-128">FunctionName</span><span class="sxs-lookup"><span data-stu-id="1078f-128">FunctionName</span></span>

<span data-ttu-id="1078f-129">Required element when **xsi:type** is "ExecuteFunction".</span><span class="sxs-lookup"><span data-stu-id="1078f-129">Required element when **xsi:type** is "ExecuteFunction".</span></span> <span data-ttu-id="1078f-130">Specifies the name of the function to execute.</span><span class="sxs-lookup"><span data-stu-id="1078f-130">Specifies the name of the function to execute.</span></span> <span data-ttu-id="1078f-131">The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span><span class="sxs-lookup"><span data-stu-id="1078f-131">The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="1078f-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1078f-132">SourceLocation</span></span>

<span data-ttu-id="1078f-133">**Xsi： type**为 "ShowTaskpane" 时必需的元素。</span><span class="sxs-lookup"><span data-stu-id="1078f-133">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="1078f-134">指定此操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="1078f-134">Specifies the source file location for this action.</span></span> <span data-ttu-id="1078f-135">**resid** 属性必须设置为 **Urls** 元素（位于 **Resources** 元素）中 **Url** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="1078f-135">The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="1078f-136">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="1078f-136">TaskpaneId</span></span>

<span data-ttu-id="1078f-137"> **xsi: type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="1078f-137">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="1078f-138">指定任务窗格容器的 ID。</span><span class="sxs-lookup"><span data-stu-id="1078f-138">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="1078f-139">具有多个“ShowTaskpane”操作时，如果想要让每个操作使用独立的窗格，则使用不同的 **TaskpaneId**。</span><span class="sxs-lookup"><span data-stu-id="1078f-139">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="1078f-140">若要让不同的操作共享同一个窗格，则使用同一个 **TaskpaneId**。</span><span class="sxs-lookup"><span data-stu-id="1078f-140">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="1078f-141">当用户选择共享同一个 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将被替换为相应的操作“SourceLocation”。</span><span class="sxs-lookup"><span data-stu-id="1078f-141">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="1078f-142">Outlook 不支持此元素。</span><span class="sxs-lookup"><span data-stu-id="1078f-142">This element is not supported in Outlook.</span></span>

<span data-ttu-id="1078f-143">下面的示例展示了两个共享同一个 **TaskpaneId** 的操作。</span><span class="sxs-lookup"><span data-stu-id="1078f-143">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="1078f-144">The following examples show two actions that use a different **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="1078f-144">The following examples show two actions that use a different **TaskpaneId**.</span></span> <span data-ttu-id="1078f-145">To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="1078f-145">To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="1078f-146">标题</span><span class="sxs-lookup"><span data-stu-id="1078f-146">Title</span></span>

<span data-ttu-id="1078f-147"> **xsi: type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="1078f-147">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="1078f-148">指定此操作任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="1078f-148">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="1078f-149">下面的示例演示使用**Title**元素的操作。</span><span class="sxs-lookup"><span data-stu-id="1078f-149">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="1078f-150">请注意，您不会直接向字符串分配**标题**。</span><span class="sxs-lookup"><span data-stu-id="1078f-150">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="1078f-151">而是为其分配一个资源 ID (resid) ，该 ID 在清单的 "**资源**" 部分中定义。</span><span class="sxs-lookup"><span data-stu-id="1078f-151">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="1078f-152">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="1078f-152">SupportsPinning</span></span>

<span data-ttu-id="1078f-153">**xsi:type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="1078f-153">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="1078f-154">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="1078f-154">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="1078f-155">添加此元素时将值设为 `true` 可以支持任务窗格固定。</span><span class="sxs-lookup"><span data-stu-id="1078f-155">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="1078f-156">这样一来，用户可以“固定”任务窗格，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="1078f-156">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="1078f-157">有关详细信息，请参阅[在 Outlook 中实现可固定的任务窗格](../../outlook/pinnable-taskpane.md)。</span><span class="sxs-lookup"><span data-stu-id="1078f-157">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1078f-158">尽管 `SupportsPinning` 在[要求集 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)中引入了此元素，但目前仅使用以下程序支持 Microsoft 365 订阅者。</span><span class="sxs-lookup"><span data-stu-id="1078f-158">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="1078f-159">Outlook 2016 或更高版本位于 Windows (内部版本7628.1000 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="1078f-159">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="1078f-160">Outlook 2016 或更高版本 Mac (build 16.13.503 or 更高版本) </span><span class="sxs-lookup"><span data-stu-id="1078f-160">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="1078f-161">新式 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="1078f-161">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
