---
title: 使用 Excel JavaScript API 处理事件
description: ''
ms.date: 09/21/2018
ms.openlocfilehash: b56d25e7e0306b4881115397d4136e63ddc03e5c
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459173"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理事件 

本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。 

## <a name="events-in-excel"></a>Excel 中的事件

每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。

| 事件 | 说明 | 支持的对象 |
|:---------------|:-------------|:-----------|
| `onAdded` | 添加对象时发生的事件。 | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | 删除对象时发生的事件。 | [**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection) |
| `onActivated` | 启用对象时发生的事件。 | [**图表**](https://docs.microsoft.com/javascript/api/excel/excel.chart)、[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**工作表**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onDeactivated` | 停用对象时发生的事件。 | [**图表**](https://docs.microsoft.com/javascript/api/excel/excel.chart)、[**ChartCollection**](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**工作表**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onCalculated` | 工作表完成计算（或集合的所有工作表都已完成）时发生的事件。 | [**WorksheetCollection**](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)、[**工作表**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) |
| `onChanged` | 更改单元格内数据时发生的事件。 | [**工作表**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**表**](https://docs.microsoft.com/javascript/api/excel/excel.table)[**、TableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.tablecollection) |
| `onDataChanged` | 更改绑定中的数据或格式时发生的事件。 | [**绑定**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSelectionChanged` | 更改活动单元格或选定范围时发生的事件。 | [**工作表**](https://docs.microsoft.com/javascript/api/excel/excel.worksheet)、[**表格**](https://docs.microsoft.com/javascript/api/excel/excel.table)、[**绑定**](https://docs.microsoft.com/javascript/api/excel/excel.binding) |
| `onSettingsChanged` | 当文档设置变更时发生的事件。 | [**SettingCollection**](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection) |

### <a name="event-triggers"></a>事件触发器

Excel 工作簿内的事件可以通过下列方式触发：

- 更改工作簿的 Excel 用户界面 (UI) 用户交互
- 更改工作簿的 Office 加载项 (JavaScript) 代码
- 更改工作簿的 VBA 加载项（宏）代码

任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。

### <a name="lifecycle-of-an-event-handler"></a>事件处理程序的生命周期

事件处理程序在加载项注册事件处理程序时创建完成，并在加载项取消注册事件处理程序或加载项关闭时销毁。事件处理程序不会暂留为 Excel 文件的一部分。

### <a name="events-and-coauthoring"></a>事件和共同创作

借助[共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。

## <a name="register-an-event-handler"></a>注册事件处理程序

下面的代码示例为 **Sample** 工作表中的 `onChanged` 事件注册事件处理程序。此代码指定 `handleDataChange` 函数应在工作表中的数据有变化时运行。

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>处理事件

如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。可以将函数设计为执行方案所需的任何操作。下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。 

```js
function handleChange(event)
{ 
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>删除事件处理程序

下面的代码示例为 **Sample** 工作表中的 `onSelectionChanged` 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{ 
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();
        
        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="enable-and-disable-events"></a>启用和禁用事件

可以通过禁用事件来提高加载项的性能。例如，您的应用可能永远不需要接收事件，也可能在执行多个实体的批量编辑时忽略事件。 

在[运行时](https://docs.microsoft.com/javascript/api/excel/excel.runtime)级别启用或禁用事件。`enableEvents` 属性判断是否会触发事件，并激活其处理程序。 

下面的代码示例演示如何打开和关闭事件。

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>另请参阅

- [使用 Excel JavaScript API 的基本编程概念](excel-add-ins-core-concepts.md)