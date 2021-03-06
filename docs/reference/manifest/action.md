---
title: 清单文件中的 Action 元素
description: 此元素指定在用户选择按钮或菜单控件时要执行的操作。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 6be1430800dea27dbd9bf78607161d88e475c145
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505407"
---
# <a name="action-element"></a><span data-ttu-id="f22ae-103">Action 元素</span><span class="sxs-lookup"><span data-stu-id="f22ae-103">Action element</span></span>

<span data-ttu-id="f22ae-104">指定在用户选择按钮或菜单控件[时](control.md#button-control)[要执行的操作](control.md#menu-dropdown-button-controls)。</span><span class="sxs-lookup"><span data-stu-id="f22ae-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="f22ae-105">属性</span><span class="sxs-lookup"><span data-stu-id="f22ae-105">Attributes</span></span>

|  <span data-ttu-id="f22ae-106">属性</span><span class="sxs-lookup"><span data-stu-id="f22ae-106">Attribute</span></span>  |  <span data-ttu-id="f22ae-107">必需</span><span class="sxs-lookup"><span data-stu-id="f22ae-107">Required</span></span>  |  <span data-ttu-id="f22ae-108">说明</span><span class="sxs-lookup"><span data-stu-id="f22ae-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f22ae-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f22ae-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="f22ae-110">是</span><span class="sxs-lookup"><span data-stu-id="f22ae-110">Yes</span></span>  | <span data-ttu-id="f22ae-111">要执行的操作类型</span><span class="sxs-lookup"><span data-stu-id="f22ae-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="f22ae-112">子元素</span><span class="sxs-lookup"><span data-stu-id="f22ae-112">Child elements</span></span>

|  <span data-ttu-id="f22ae-113">元素</span><span class="sxs-lookup"><span data-stu-id="f22ae-113">Element</span></span> |  <span data-ttu-id="f22ae-114">说明</span><span class="sxs-lookup"><span data-stu-id="f22ae-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="f22ae-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="f22ae-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="f22ae-116">指定要执行的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="f22ae-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="f22ae-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f22ae-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="f22ae-118">指定该操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="f22ae-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="f22ae-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="f22ae-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="f22ae-120">指定任务窗格容器的 ID。</span><span class="sxs-lookup"><span data-stu-id="f22ae-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="f22ae-121">Title</span><span class="sxs-lookup"><span data-stu-id="f22ae-121">Title</span></span>](#title) | <span data-ttu-id="f22ae-122">指定任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="f22ae-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="f22ae-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="f22ae-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="f22ae-124">指定任务窗格支持固定，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="f22ae-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|

## <a name="xsitype"></a><span data-ttu-id="f22ae-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f22ae-125">xsi:type</span></span>

<span data-ttu-id="f22ae-p101">此属性指定当用户选择按钮时所执行的操作类型。可取值如下：</span><span class="sxs-lookup"><span data-stu-id="f22ae-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> <span data-ttu-id="f22ae-128">注册 [邮箱](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) 和 [项目](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) 事件在 **xsi：type** 为 时不可用 `ExecuteFunction` 。</span><span class="sxs-lookup"><span data-stu-id="f22ae-128">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available when **xsi:type** is `ExecuteFunction`.</span></span>

## <a name="functionname"></a><span data-ttu-id="f22ae-129">FunctionName</span><span class="sxs-lookup"><span data-stu-id="f22ae-129">FunctionName</span></span>

<span data-ttu-id="f22ae-p102">**xsi:type** 为“ExecuteFunction”时的必需元素。指定要执行的函数的名称。函数包含在 [FunctionFile](functionfile.md) 元素指定的文件中。</span><span class="sxs-lookup"><span data-stu-id="f22ae-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="f22ae-133">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f22ae-133">SourceLocation</span></span>

<span data-ttu-id="f22ae-134">**xsi：type** 为"ShowTaskpane"时所需的元素。</span><span class="sxs-lookup"><span data-stu-id="f22ae-134">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="f22ae-135">指定该操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="f22ae-135">Specifies the source file location for this action.</span></span> <span data-ttu-id="f22ae-136">**resid** 属性不能超过 32 个字符，并且必须设置为 Resources 元素 **中 Urls** 元素中 **Url** 元素 **的 id** [属性值。](resources.md)</span><span class="sxs-lookup"><span data-stu-id="f22ae-136">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="f22ae-137">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="f22ae-137">TaskpaneId</span></span>

<span data-ttu-id="f22ae-p104">可选元素，当 **xsi: type** 是“ShowTaskpane”时。指定任务窗格容器的 ID。具有多个“ShowTaskpane”操作时，如果想要对每个操作使用独立的窗格，则使用不同的 **TaskpaneId**。为共享相同窗格的不同操作使用同一 **TaskpaneId** 当用户选择共享同一 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将被替换为相应的操作“SourceLocation”</span><span class="sxs-lookup"><span data-stu-id="f22ae-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="f22ae-143">Outlook 不支持此元素。</span><span class="sxs-lookup"><span data-stu-id="f22ae-143">This element is not supported in Outlook.</span></span>

<span data-ttu-id="f22ae-144">下面的示例展示了两个共享同一个 **TaskpaneId** 的操作。</span><span class="sxs-lookup"><span data-stu-id="f22ae-144">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="f22ae-p105">下面的示例展示了两个使用不同 **TaskpaneId** 的操作。若要查看上下文中的这些示例，请参阅 [简单的外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)。</span><span class="sxs-lookup"><span data-stu-id="f22ae-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="f22ae-147">标题</span><span class="sxs-lookup"><span data-stu-id="f22ae-147">Title</span></span>

<span data-ttu-id="f22ae-148">**xsi: type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="f22ae-148">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="f22ae-149">指定此操作任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="f22ae-149">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="f22ae-150">以下示例演示使用 **Title** 元素的操作。</span><span class="sxs-lookup"><span data-stu-id="f22ae-150">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="f22ae-151">请注意，不要直接将 **Title** 分配给字符串。</span><span class="sxs-lookup"><span data-stu-id="f22ae-151">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="f22ae-152">相反，你可以为其分配 (一) ID，该 ID 在清单的 **"资源** "部分中定义，并且不能超过 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="f22ae-152">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="f22ae-153">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="f22ae-153">SupportsPinning</span></span>

<span data-ttu-id="f22ae-154">**xsi:type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="f22ae-154">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="f22ae-155">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="f22ae-155">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="f22ae-156">添加此元素时将值设为 `true` 可以支持任务窗格固定。</span><span class="sxs-lookup"><span data-stu-id="f22ae-156">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="f22ae-157">这样一来，用户可以“固定”任务窗格，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="f22ae-157">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="f22ae-158">有关详细信息，请参阅[在 Outlook 中实现可固定的任务窗格](../../outlook/pinnable-taskpane.md)。</span><span class="sxs-lookup"><span data-stu-id="f22ae-158">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f22ae-159">尽管 `SupportsPinning` 元素是在要求集 [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)中引入的，但它当前仅支持使用以下内容的 Microsoft 365 订阅者。</span><span class="sxs-lookup"><span data-stu-id="f22ae-159">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
>
> - <span data-ttu-id="f22ae-160">Windows 版 Outlook 2016 或更高版本 (版本 7628.1000 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="f22ae-160">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="f22ae-161">Mac 版 Outlook 2016 或 (版本 16.13.503 或更高版本) </span><span class="sxs-lookup"><span data-stu-id="f22ae-161">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="f22ae-162">新式 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="f22ae-162">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
