---
title: 清单文件中的 Action 元素
description: 此元素指定在用户选择按钮或菜单控件时要执行的操作。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: e345d0a1682e0125373a309e1e56eb2d6298ac7d
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771408"
---
# <a name="action-element"></a>Action 元素

指定当用户选择按钮或菜单控件[时](control.md#button-control)[要执行的操作](control.md#menu-dropdown-button-controls)。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 要执行的操作类型|

## <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [FunctionName](#functionname) |    指定要执行的函数的名称。 |
|  [SourceLocation](#sourcelocation) |    指定该操作的源文件位置。 |
|  [TaskpaneId](#taskpaneid) | 指定任务窗格容器的 ID。|
|  [Title](#title) | 指定任务窗格的自定义标题。|
|  [SupportsPinning](#supportspinning) | 指定任务窗格支持固定，即使用户选择其他对象，任务窗格也可以继续处于打开状态。|
  

## <a name="xsitype"></a>xsi:type

此属性指定当用户选择按钮时所执行的操作类型。可取值如下：

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a>FunctionName

**xsi:type** 为“ExecuteFunction”时的必需元素。指定要执行的函数的名称。函数包含在 [FunctionFile](functionfile.md) 元素指定的文件中。

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

**xsi：type** 为"ShowTaskpane"时所需的元素。 指定该操作的源文件位置。 **resid** 属性不能超过 32 个字符，必须设置为 Resources 元素中 **Url 元素** 中 **Url** 元素 **的 id**[属性值。](resources.md)

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

可选元素，当 **xsi: type** 是“ShowTaskpane”时。指定任务窗格容器的 ID。具有多个“ShowTaskpane”操作时，如果想要对每个操作使用独立的窗格，则使用不同的 **TaskpaneId**。为共享相同窗格的不同操作使用同一 **TaskpaneId** 当用户选择共享同一 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将被替换为相应的操作“SourceLocation”

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

下面的示例展示了两个使用不同 **TaskpaneId** 的操作。若要查看上下文中的这些示例，请参阅 [简单的外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml)。

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

下面的示例演示使用 **Title** 元素的操作。 请注意，不要直接将 **Title** 分配给字符串。 相反，你可以为其分配 (id) ，该 ID 在清单的 **"资源** "部分中定义，并且不能超过 32 个字符。

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
> 尽管 `SupportsPinning` 元素是在要求集 [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)中引入的，但当前仅支持使用以下内容的 Microsoft 365 订阅者。
> - Windows 版 Outlook 2016 或 (版本 7628.1000 或更高版本) 
> - Mac 版 Outlook 2016 或 (版本 16.13.503 或更高版本) 
> - 新式 Outlook 网页版

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
