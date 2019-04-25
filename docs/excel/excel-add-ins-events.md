---
title: 使用 Excel JavaScript API 处理事件
description: ''
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: 7f05263f5220c2d60d0cebcfc686e1fed3f07900
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449265"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理事件

本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。 

## <a name="events-in-excel"></a>Excel 中的事件

每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。 使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 下列事件暂不受支持。

| 事件 | 说明 | 支持的对象 |
|:---------------|:-------------|:-----------|
| `onActivated` | 激活对象时发生。 | [**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onAdded` | 添加对象时发生。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onCalculated` | 工作表完成计算（或集合的所有工作表都已完成）时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onChanged` | 更改单元格内的数据时发生。 | [**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet) |
| `onDataChanged` | 当绑定内的数据或格式变化时发生。 | [**Binding**](/javascript/api/excel/excel.binding) |
| `onDeactivated` | 停用对象时发生。 | [**Chart**](/javascript/api/excel/excel.chart)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onDeleted` | 删除对象时发生。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onSelectionChanged` | 更改活动单元格或选定范围时发生。 | [**Binding**](/javascript/api/excel/excel.binding)、[**Table**](/javascript/api/excel/excel.table)、[**Worksheet**](/javascript/api/excel/excel.worksheet) |
| `onSettingsChanged` | 当文档中的设置变化时发生。 | [**SettingCollection**](/javascript/api/excel/excel.settingcollection) |

### <a name="events-in-preview"></a>预览版中的事件

> [!NOTE]
> 以下事件当前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| 事件 | 说明 | 支持的对象 |
|:---------------|:-------------|:-----------|
| `onActivated` | 当激活形状时发生。 | [**Shape**](/javascript/api/excel/excel.shape)|
| `onAdded` | 在工作簿中添加新表格时发生。 | [**TableCollection**](/javascript/api/excel/excel.tablecollection)|
| `onAutoSaveSettingChanged` | 在工作簿上更改 `autoSave` 设置时发生。 | [**Workbook**](/javascript/api/excel/excel.workbook) |
| `onChanged` | 在更改工作簿中的任何工作表时发生。 | [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection)|
| `onDeactivated` | 当停用形状时发生。 | [**Shape**](/javascript/api/excel/excel.shape)|
| `onDeleted` | 在工作簿中删除指定的表格时发生。 | [**TableCollection**](/javascript/api/excel/excel.tablecollection)|
| `onFiltered` | 在对象上应用筛选器时发生。 | [**Table**](/javascript/api/excel/excel.table)、[**TableCollection**](/javascript/api/excel/excel.tablecollection)、[**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onFormatChanged` | 在工作表上的格式变化时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |
| `onSelectionChanged` | 在任何工作表上更改选择时发生。 | [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection) |

### <a name="event-triggers"></a>事件触发器

Excel 工作簿内的事件可以通过下列方式触发：

- 更改工作簿的 Excel 用户界面 (UI) 用户交互
- 更改工作簿的 Office 加载项 (JavaScript) 代码
- 更改工作簿的 VBA 加载项（宏）代码

任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。

### <a name="lifecycle-of-an-event-handler"></a>事件处理程序的生命周期

当加载项注册事件处理程序时，将创建事件处理程序。 当加载项取消注册事件处理程序或者刷新、重新加载或关闭加载项时，将销毁事件处理程序。 事件处理程序不会作为 Excel 文件的一部分保留，也不会在与 Excel Online 的会话之间保留。

> [!CAUTION]
> 删除了注册事件的对象（例如，注册 `onChanged` 事件的表）时，事件处理程序不再触发但会保留在内存中，直到加载项或 Excel 会话刷新或关闭为止。

### <a name="events-and-coauthoring"></a>事件和共同创作

借助[共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。

## <a name="register-an-event-handler"></a>注册事件处理程序

下面的代码示例为 `onChanged` 工作表中的 **** 事件注册事件处理程序。 此代码指定 `handleDataChange` 函数应在工作表中的数据有变化时运行。

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

如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。 可以将函数设计为执行方案所需的任何操作。 下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。

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

下面的代码示例为 `onSelectionChanged` 工作表中的 **** 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。 它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。

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

可以通过禁用事件来改进加载项性能。
例如，你的应用可能永远不需要接收事件，也可能在执行多个实体的批量编辑时忽略事件。

启用和禁用事件是在[运行时](/javascript/api/excel/excel.runtime)级别进行的。
`enableEvents` 属性确定是否触发事件并激活其处理程序。

以下代码示例展示了如何打开和关闭事件。

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

- [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)
