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
# <a name="action-element"></a>Action 元素

指定当用户选择[按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件时要执行的操作。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 要执行的操作类型|

## <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [FunctionName](#functionname) |    指定要执行的函数的名称。 |
|  [SourceLocation](#sourcelocation) |    指定该操作的源文件位置。 |
|  [TaskpaneId](#taskpaneid) | 指定任务窗格容器的 ID。|
|  [Title](#title) | 指定任务窗格的自定义标题。|
|  [SupportsPinning](#supportspinning) | 指定任务窗格支持固定，即使用户选择其他对象，任务窗格也可以继续处于打开状态。|
  

## <a name="xsitype"></a>xsi:type

This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

**Xsi： type**为 "ShowTaskpane" 时必需的元素。 指定此操作的源文件位置。 **resid** 属性必须设置为 **Urls** 元素（位于 **Resources** 元素）中 **Url** 元素的 [id](resources.md) 属性的值。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

 **xsi: type** 是“ShowTaskpane”时的可选元素。 指定任务窗格容器的 ID。 具有多个“ShowTaskpane”操作时，如果想要让每个操作使用独立的窗格，则使用不同的 **TaskpaneId**。 若要让不同的操作共享同一个窗格，则使用同一个 **TaskpaneId**。 当用户选择共享同一个 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将被替换为相应的操作“SourceLocation”。

> [!NOTE]
> Outlook 不支持此元素。

下面的示例展示了两个共享同一个 **TaskpaneId** 的操作。

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

The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).

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

## <a name="title"></a>标题

 **xsi: type** 是“ShowTaskpane”时的可选元素。 指定此操作任务窗格的自定义标题。

下面的示例演示使用**Title**元素的操作。 请注意，您不会直接向字符串分配**标题**。 而是为其分配一个资源 ID (resid) ，该 ID 在清单的 "**资源**" 部分中定义。

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

## <a name="supportspinning"></a>SupportsPinning

**xsi:type** 是“ShowTaskpane”时的可选元素。 包含 [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。 添加此元素时将值设为 `true` 可以支持任务窗格固定。 这样一来，用户可以“固定”任务窗格，即使用户选择其他对象，任务窗格也可以继续处于打开状态。 有关详细信息，请参阅[在 Outlook 中实现可固定的任务窗格](../../outlook/pinnable-taskpane.md)。

> [!IMPORTANT]
> 尽管 `SupportsPinning` 在[要求集 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)中引入了此元素，但目前仅使用以下程序支持 Microsoft 365 订阅者。
> - Outlook 2016 或更高版本位于 Windows (内部版本7628.1000 或更高版本) 
> - Outlook 2016 或更高版本 Mac (build 16.13.503 or 更高版本) 
> - 新式 Outlook 网页版

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
