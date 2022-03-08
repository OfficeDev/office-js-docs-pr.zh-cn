---
title: 清单文件中的 Action 元素
description: 此元素指定在用户选择按钮或菜单控件时要执行的操作。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21c8f9a6345641f23aad70efed67c9c45f72a1c8
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340412"
---
# <a name="action-element"></a>Action 元素

指定在用户选择"按钮"或"菜单"控件[时](control-button.md)[要执行的操作](control-menu.md)。

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- 当父 **VersionOverrides** 的类型为 Taskpane [1.0 时，AddinCommands](../requirement-sets/add-in-commands-requirement-sets.md) 1.1。
- [当父](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) **VersionOverrides** 类型为 Mail 1.0 时，邮箱 1.3。
- [当父](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) **VersionOverrides** 类型为 Mail 1.1 时，邮箱 1.5。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 要执行的操作类型|

## <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [FunctionName](#functionname) |    指定要执行的函数的名称。 |
|  [SourceLocation](#sourcelocation) |    指定该操作的源文件位置。 |
|  [TaskpaneId](#taskpaneid) | 指定任务窗格容器的 ID。 在加载项Outlook不支持。|
|  [Title](#title) | 指定任务窗格的自定义标题。 在加载项Outlook不支持。|
|  [SupportsPinning](#supportspinning) | 指定任务窗格支持固定，即使用户选择其他对象，任务窗格也可以继续处于打开状态。|

## <a name="xsitype"></a>xsi:type

此属性指定当用户选择按钮时所执行的操作类型。可取值如下：

- `ExecuteFunction`
- `ShowTaskpane`

> [!IMPORTANT]
> 当 [](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) **xsi：type 为 时，注册邮箱和项目事件** 不可用。[](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) `ExecuteFunction`

## <a name="functionname"></a>FunctionName

**xsi：type** 为 时必需的元素`ExecuteFunction`。 指定要执行的函数的名称。 函数包含在 [FunctionFile](functionfile.md) 元素指定的文件中。

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation

**xsi：type** 为 时必需的元素`ShowTaskpane`。 指定该操作的源文件位置。 **resid** 属性不能超过 32 个字符，并且必须设置为 **Urls** 元素（位于 [Resources](resources.md) 元素）中 Url **元素** 的 **id** 属性的值。

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a>TaskpaneId

**xsi：type** 为 的可选元素`ShowTaskpane`。 指定任务窗格容器的 ID。 当有多个操作时 `ShowTaskpane` ，如果你需要每个操作的独立窗格，请使用不同的 **TaskpaneId** 。 若要让不同的操作共享同一个窗格，则使用同一个 **TaskpaneId**。 当用户选择共享同一 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将替换为相应的操作 `SourceLocation`。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

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

**xsi：type** 为 的可选元素`ShowTaskpane`。 指定此操作任务窗格的自定义标题。

**外接程序类型：** 任务窗格

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

> [!NOTE]
> 此子元素在加载项Outlook支持。

以下示例演示使用 **Title** 元素的操作。 请注意，不要直接将 **Title** 分配给字符串。 相反，你可以为其分配 (resid) ，该 ID 在清单的"资源"部分中定义，且不能超过  32 个字符。

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

**xsi：type** 为 的可选元素`ShowTaskpane`。 包含 [VersionOverrides](versionoverrides.md) 元素的 **xsi：type** 属性值必须为 `VersionOverridesV1_1`。 添加此元素时将值设为 `true` 可以支持任务窗格固定。 这样一来，用户可以“固定”任务窗格，即使用户选择其他对象，任务窗格也可以继续处于打开状态。 有关详细信息，请参阅[在 Outlook 中实现可固定的任务窗格](../../outlook/pinnable-taskpane.md)。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [Mailbox 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

> [!IMPORTANT]
> 尽管 **SupportsPinning** 元素是在要求集 [1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) 中引入的，但当前仅支持以下Microsoft 365订阅者：
>
> - Outlook 2016版本 7628.1000 或Windows (版本 7628.1000 或更高版本) 
> - Outlook 2016版本 16.13.503 (更高版本的 Mac 版本或更高版本) 
> - 新式 Outlook 网页版

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
