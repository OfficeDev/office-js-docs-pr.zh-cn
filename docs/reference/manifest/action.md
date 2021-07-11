---
title: 清单文件中的 Action 元素
description: 此元素指定在用户选择按钮或菜单控件时要执行的操作。
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 1ec2623ad5dbb07677735b7bcb1e39612e56984c
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348697"
---
# <a name="action-element"></a><span data-ttu-id="fb58c-103">Action 元素</span><span class="sxs-lookup"><span data-stu-id="fb58c-103">Action element</span></span>

<span data-ttu-id="fb58c-104">指定在用户选择"按钮"或"菜单"控件[时](control.md#button-control)[要执行的操作](control.md#menu-dropdown-button-controls)。</span><span class="sxs-lookup"><span data-stu-id="fb58c-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="fb58c-105">属性</span><span class="sxs-lookup"><span data-stu-id="fb58c-105">Attributes</span></span>

|  <span data-ttu-id="fb58c-106">属性</span><span class="sxs-lookup"><span data-stu-id="fb58c-106">Attribute</span></span>  |  <span data-ttu-id="fb58c-107">必需</span><span class="sxs-lookup"><span data-stu-id="fb58c-107">Required</span></span>  |  <span data-ttu-id="fb58c-108">说明</span><span class="sxs-lookup"><span data-stu-id="fb58c-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fb58c-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fb58c-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="fb58c-110">是</span><span class="sxs-lookup"><span data-stu-id="fb58c-110">Yes</span></span>  | <span data-ttu-id="fb58c-111">要执行的操作类型</span><span class="sxs-lookup"><span data-stu-id="fb58c-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="fb58c-112">子元素</span><span class="sxs-lookup"><span data-stu-id="fb58c-112">Child elements</span></span>

|  <span data-ttu-id="fb58c-113">元素</span><span class="sxs-lookup"><span data-stu-id="fb58c-113">Element</span></span> |  <span data-ttu-id="fb58c-114">说明</span><span class="sxs-lookup"><span data-stu-id="fb58c-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="fb58c-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="fb58c-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="fb58c-116">指定要执行的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="fb58c-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="fb58c-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="fb58c-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="fb58c-118">指定该操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="fb58c-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="fb58c-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="fb58c-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="fb58c-120">指定任务窗格容器的 ID。</span><span class="sxs-lookup"><span data-stu-id="fb58c-120">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="fb58c-121">在加载项Outlook不支持。</span><span class="sxs-lookup"><span data-stu-id="fb58c-121">Not supported in Outlook add-ins.</span></span>|
|  [<span data-ttu-id="fb58c-122">Title</span><span class="sxs-lookup"><span data-stu-id="fb58c-122">Title</span></span>](#title) | <span data-ttu-id="fb58c-123">指定任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="fb58c-123">Specifies the custom title for the task pane.</span></span> <span data-ttu-id="fb58c-124">在加载项Outlook不支持。</span><span class="sxs-lookup"><span data-stu-id="fb58c-124">Not supported in Outlook add-ins.</span></span>|
|  [<span data-ttu-id="fb58c-125">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="fb58c-125">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="fb58c-126">指定任务窗格支持固定，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="fb58c-126">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|

## <a name="xsitype"></a><span data-ttu-id="fb58c-127">xsi:type</span><span class="sxs-lookup"><span data-stu-id="fb58c-127">xsi:type</span></span>

<span data-ttu-id="fb58c-p103">此属性指定当用户选择按钮时所执行的操作类型。可取值如下：</span><span class="sxs-lookup"><span data-stu-id="fb58c-p103">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> <span data-ttu-id="fb58c-130">当 [](../objectmodel/preview-requirement-set/office.context.mailbox.md#events)**xsi：type** 为 时，注册邮箱和项目事件不可用 [](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) `ExecuteFunction` 。</span><span class="sxs-lookup"><span data-stu-id="fb58c-130">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available when **xsi:type** is `ExecuteFunction`.</span></span>

## <a name="functionname"></a><span data-ttu-id="fb58c-131">FunctionName</span><span class="sxs-lookup"><span data-stu-id="fb58c-131">FunctionName</span></span>

<span data-ttu-id="fb58c-p104">**xsi:type** 为“ExecuteFunction”时的必需元素。指定要执行的函数的名称。函数包含在 [FunctionFile](functionfile.md) 元素指定的文件中。</span><span class="sxs-lookup"><span data-stu-id="fb58c-p104">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="fb58c-135">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="fb58c-135">SourceLocation</span></span>

<span data-ttu-id="fb58c-136">**xsi：type** 为"ShowTaskpane"时必需元素。</span><span class="sxs-lookup"><span data-stu-id="fb58c-136">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="fb58c-137">指定该操作的源文件位置。</span><span class="sxs-lookup"><span data-stu-id="fb58c-137">Specifies the source file location for this action.</span></span> <span data-ttu-id="fb58c-138">**resid** 属性不能超过 32 个字符，并且必须设置为 **Urls** 元素（位于 [Resources](resources.md)元素）中 Url **元素** 的 **id** 属性的值。</span><span class="sxs-lookup"><span data-stu-id="fb58c-138">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="fb58c-139">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="fb58c-139">TaskpaneId</span></span>

<span data-ttu-id="fb58c-p106">可选元素，当 **xsi: type** 是“ShowTaskpane”时。指定任务窗格容器的 ID。具有多个“ShowTaskpane”操作时，如果想要对每个操作使用独立的窗格，则使用不同的 **TaskpaneId**。为共享相同窗格的不同操作使用同一 **TaskpaneId** 当用户选择共享同一 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将被替换为相应的操作“SourceLocation”</span><span class="sxs-lookup"><span data-stu-id="fb58c-p106">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="fb58c-145">Outlook 不支持此元素。</span><span class="sxs-lookup"><span data-stu-id="fb58c-145">This element is not supported in Outlook.</span></span>

<span data-ttu-id="fb58c-146">下面的示例展示了两个共享同一个 **TaskpaneId** 的操作。</span><span class="sxs-lookup"><span data-stu-id="fb58c-146">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="fb58c-p107">下面的示例展示了两个使用不同 **TaskpaneId** 的操作。若要查看上下文中的这些示例，请参阅 [简单的外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)。</span><span class="sxs-lookup"><span data-stu-id="fb58c-p107">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="fb58c-149">标题</span><span class="sxs-lookup"><span data-stu-id="fb58c-149">Title</span></span>

<span data-ttu-id="fb58c-150">**xsi: type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="fb58c-150">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="fb58c-151">指定此操作任务窗格的自定义标题。</span><span class="sxs-lookup"><span data-stu-id="fb58c-151">Specifies the custom title for the task pane for this action.</span></span>

> [!NOTE]
> <span data-ttu-id="fb58c-152">此子元素在加载项中Outlook支持。</span><span class="sxs-lookup"><span data-stu-id="fb58c-152">This child element is not supported in Outlook add-ins.</span></span>

<span data-ttu-id="fb58c-153">以下示例演示使用 **Title** 元素的操作。</span><span class="sxs-lookup"><span data-stu-id="fb58c-153">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="fb58c-154">请注意，不要直接将 **Title** 分配给字符串。</span><span class="sxs-lookup"><span data-stu-id="fb58c-154">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="fb58c-155">相反，你可以为其分配 (resid) ，该 ID 在清单的"资源"部分中定义，且不能超过 32 个字符。</span><span class="sxs-lookup"><span data-stu-id="fb58c-155">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

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

## <a name="supportspinning"></a><span data-ttu-id="fb58c-156">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="fb58c-156">SupportsPinning</span></span>

<span data-ttu-id="fb58c-157">**xsi:type** 是“ShowTaskpane”时的可选元素。</span><span class="sxs-lookup"><span data-stu-id="fb58c-157">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="fb58c-158">包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。</span><span class="sxs-lookup"><span data-stu-id="fb58c-158">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="fb58c-159">添加此元素时将值设为 `true` 可以支持任务窗格固定。</span><span class="sxs-lookup"><span data-stu-id="fb58c-159">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="fb58c-160">这样一来，用户可以“固定”任务窗格，即使用户选择其他对象，任务窗格也可以继续处于打开状态。</span><span class="sxs-lookup"><span data-stu-id="fb58c-160">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="fb58c-161">有关详细信息，请参阅[在 Outlook 中实现可固定的任务窗格](../../outlook/pinnable-taskpane.md)。</span><span class="sxs-lookup"><span data-stu-id="fb58c-161">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fb58c-162">尽管 `SupportsPinning` 元素是在要求集[1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)中引入的，但当前仅支持以下Microsoft 365订阅者：</span><span class="sxs-lookup"><span data-stu-id="fb58c-162">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following:</span></span>
>
> - <span data-ttu-id="fb58c-163">Outlook 2016版本 7628.1000 或 Windows (更高版本) </span><span class="sxs-lookup"><span data-stu-id="fb58c-163">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="fb58c-164">Outlook 2016版本 16.13.503 (更高版本的 Mac 版本或更高版本) </span><span class="sxs-lookup"><span data-stu-id="fb58c-164">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="fb58c-165">新式 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="fb58c-165">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
