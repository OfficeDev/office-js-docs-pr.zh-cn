---
title: 清单文件中的 Action 元素
description: ''
ms.date: 11/14/2018
localization_priority: Normal
ms.openlocfilehash: 589a4af94c7abbcf61cd7a5210d5df29ba8a3a4e
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635893"
---
# <a name="action-element"></a><span data-ttu-id="d0dcb-102">Action 元素</span><span class="sxs-lookup"><span data-stu-id="d0dcb-102">Action element</span></span>

<span data-ttu-id="d0dcb-103">指定用户选择 [按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件时将执行的操作。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-103">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="d0dcb-104">属性</span><span class="sxs-lookup"><span data-stu-id="d0dcb-104">Attributes</span></span>

|  <span data-ttu-id="d0dcb-105">属性</span><span class="sxs-lookup"><span data-stu-id="d0dcb-105">Attribute</span></span>  |  <span data-ttu-id="d0dcb-106">必需</span><span class="sxs-lookup"><span data-stu-id="d0dcb-106">Required</span></span>  |  <span data-ttu-id="d0dcb-107">说明</span><span class="sxs-lookup"><span data-stu-id="d0dcb-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d0dcb-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="d0dcb-108">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="d0dcb-109">是</span><span class="sxs-lookup"><span data-stu-id="d0dcb-109">Yes</span></span>  | <span data-ttu-id="d0dcb-110">要执行的操作类型</span><span class="sxs-lookup"><span data-stu-id="d0dcb-110">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="d0dcb-111">子元素</span><span class="sxs-lookup"><span data-stu-id="d0dcb-111">Child elements</span></span>

|  <span data-ttu-id="d0dcb-112">元素</span><span class="sxs-lookup"><span data-stu-id="d0dcb-112">Element</span></span> |  <span data-ttu-id="d0dcb-113">说明</span><span class="sxs-lookup"><span data-stu-id="d0dcb-113">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="d0dcb-114">FunctionName</span><span class="sxs-lookup"><span data-stu-id="d0dcb-114">FunctionName</span></span>](#functionname) |    <span data-ttu-id="d0dcb-115">指定要执行的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-115">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="d0dcb-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d0dcb-116">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="d0dcb-117">指定该操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-117">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="d0dcb-118"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="d0dcb-118"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="d0dcb-119">指定任务窗格容器的 ID。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-119">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="d0dcb-120"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="d0dcb-120"> [Title](#title)</span></span> | <span data-ttu-id="d0dcb-121">指定任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-121">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="d0dcb-122"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="d0dcb-122"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="d0dcb-123">指定任务窗格支持固定，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-123">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="d0dcb-124">xsi:type</span><span class="sxs-lookup"><span data-stu-id="d0dcb-124">xsi:type</span></span>

<span data-ttu-id="d0dcb-p101">此属性指定当用户选择按钮时所执行的操作类型。可取值如下：</span><span class="sxs-lookup"><span data-stu-id="d0dcb-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="d0dcb-127">FunctionName</span><span class="sxs-lookup"><span data-stu-id="d0dcb-127">FunctionName</span></span>

<span data-ttu-id="d0dcb-p102">**xsi:type** 为“ExecuteFunction”时的必需元素。指定要执行的函数的名称。函数包含在 [FunctionFile](functionfile.md) 元素指定的文件中。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="d0dcb-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="d0dcb-131">SourceLocation</span></span>

<span data-ttu-id="d0dcb-p103">**xsi:type** 为 ShowTaskpane 时的必需元素。指定此操作的源文件位置。 **resid** 属性必须设置为 **Urls** 元素（位于 **Resources** 元素）中 **Url** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="d0dcb-135">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="d0dcb-135">TaskpaneId</span></span>

<span data-ttu-id="d0dcb-136"> *\*xsi:type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-136">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="d0dcb-137">指定任务窗格容器的 ID。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-137">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="d0dcb-138">具有多个“ShowTaskpane”操作时，如果想要让每个操作使用独立的窗格，则使用不同的 **TaskpaneId**。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-138">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="d0dcb-139">若要让不同的操作共享同一个窗格，则使用同一个 **TaskpaneId**。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-139">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="d0dcb-140">当用户选择共享同一个 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将被替换为相应的操作“SourceLocation”。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-140">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="d0dcb-141">Outlook 不支持此元素。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-141">This element is not supported in Outlook.</span></span>

<span data-ttu-id="d0dcb-142">下面的示例展示了两个共享同一个 **TaskpaneId** 的操作。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-142">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="d0dcb-p105">下面的示例展示了两个使用不同 **TaskpaneId** 的操作。若要查看上下文中的这些示例，请参阅[简单的外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="d0dcb-145">标题</span><span class="sxs-lookup"><span data-stu-id="d0dcb-145">Title</span></span>

<span data-ttu-id="d0dcb-146"> *\*xsi:type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-146">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="d0dcb-147">指定此操作任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-147">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="d0dcb-148">下面的示例展示了两个使用 **Title** 元素的不同操作。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-148">The following examples show two different actions that use the **Title** element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
```

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
```

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
```

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
```

## <a name="supportspinning"></a><span data-ttu-id="d0dcb-149">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="d0dcb-149">SupportsPinning</span></span>

<span data-ttu-id="d0dcb-150">**xsi:type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-150">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="d0dcb-151">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-151">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="d0dcb-152">添加此元素时将值设为 `true` 可以支持任务窗格固定。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-152">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="d0dcb-153">这样一来，用户可以“固定”任务窗格，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-153">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="d0dcb-154">有关详细信息，请参阅[在 Outlook 中实现可固定的任务窗格](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane)。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-154">For more information, see [Implement a pinnable task pane in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="d0dcb-155">SupportsPinning 目前仅支持 Windows （构建 7628.1000 或更高版本） 的 Outlook 2016 和 Outlook for Mac 的 2016年 (生成 16.13.503 或更高版本)。</span><span class="sxs-lookup"><span data-stu-id="d0dcb-155">SupportsPinning is currently only supported by Outlook 2016 for Windows (build 7628.1000 or later) and Outlook 2016 for Mac (build 16.13.503 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
