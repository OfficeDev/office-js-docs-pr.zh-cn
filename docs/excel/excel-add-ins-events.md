---
title: 使用 Excel JavaScript API 处理事件
description: JavaScript 对象Excel列表。 这包括有关使用事件处理程序和相关模式的信息。
ms.date: 06/04/2021
localization_priority: Normal
ms.openlocfilehash: 0a13508c501d30d74f1d21e15cf8f4e09b3f1c6a
ms.sourcegitcommit: 5a151d4df81e5640363774406d0f329d6a0d3db8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/09/2021
ms.locfileid: "52853974"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理事件

本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。

## <a name="events-in-excel"></a>Excel 中的事件

每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。 使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 下列事件暂不受支持。

| 事件 | 说明 | 支持的对象 |
|:---------------|:-------------|:-----------|
| `onActivated` | 激活对象时发生。 | [**Chart**](/javascript/api/excel/excel.chart#onactivated)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#onactivated)、[**Shape**](/javascript/api/excel/excel.shape#onactivated)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onactivated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onactivated) |
| `onAdded` | 当向集合中添加对象时发生。 | [](/javascript/api/excel/excel.chartcollection#onadded)ChartCollection、CommentCollection、TableCollection、WorksheetCollection [](/javascript/api/excel/excel.commentcollection#onadded) [](/javascript/api/excel/excel.tablecollection#onadded) [](/javascript/api/excel/excel.worksheetcollection#onadded) |
| `onAutoSaveSettingChanged` | 在工作簿上更改 `autoSave` 设置时发生。 | [**Workbook**](/javascript/api/excel/excel.workbook#onautosavesettingchanged) |
| `onCalculated` | 工作表完成计算（或集合的所有工作表都已完成）时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#oncalculated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncalculated) |
| `onChanged` | 在单个单元格或批注的数据发生更改时发生。 | [](/javascript/api/excel/excel.worksheet#onchanged)CommentCollection、Table、TableCollection、Worksheet、WorksheetCollection [](/javascript/api/excel/excel.commentcollection#onchanged) [](/javascript/api/excel/excel.table#onchanged) [](/javascript/api/excel/excel.tablecollection#onchanged) [](/javascript/api/excel/excel.worksheetcollection#onchanged) |
| `onColumnSorted` | 在已对一个或多个列进行排序时发生。 这是从左到右排序操作的结果。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#oncolumnsorted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted) |
| `onDataChanged` | 当绑定内的数据或格式变化时发生。 | [**Binding**](/javascript/api/excel/excel.binding#ondatachanged) |
| `onDeactivated` | 停用对象时发生。 | [**Chart**](/javascript/api/excel/excel.chart#ondeactivated)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#ondeactivated)、[**Shape**](/javascript/api/excel/excel.shape#ondeactivated)、[**Worksheet**](/javascript/api/excel/excel.worksheet#ondeactivated)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#ondeactivated) |
| `onDeleted` | 当从集合中删除对象时发生。 | [](/javascript/api/excel/excel.chartcollection#ondeleted)ChartCollection、CommentCollection、TableCollection、WorksheetCollection [](/javascript/api/excel/excel.commentcollection#ondeleted) [](/javascript/api/excel/excel.tablecollection#ondeleted) [](/javascript/api/excel/excel.worksheetcollection#ondeleted) |
| `onFormatChanged` | 在工作表上的格式变化时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onformatchanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onformatchanged) |
| `onRowSorted` | 在已对一个或多个行进行排序时发生。 这是从上到下排序操作的结果。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onrowsorted)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowsorted) |
| `onSelectionChanged` | 当活动单元格或选定范围更改时发生。 | [](/javascript/api/excel/excel.binding#onselectionchanged) [](/javascript/api/excel/excel.table#onselectionchanged)Binding、Table、Workbook、Worksheet、WorksheetCollection [](/javascript/api/excel/excel.workbook#onselectionchanged) [](/javascript/api/excel/excel.worksheet#onselectionchanged) [](/javascript/api/excel/excel.worksheetcollection#onselectionchanged) |
| `onRowHiddenChanged` | 在特定工作表上的行隐藏状态更改时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged) |
| `onSettingsChanged` | 当文档中的设置变化时发生。 | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#onsettingschanged) |
| `onSingleClicked` | 在工作表中进行左键单击/点击操作时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onsingleclicked)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onsingleclicked) |

### <a name="events-in-preview"></a>预览版中的事件

> [!NOTE]
> 以下事件当前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| 事件 | 说明 | 支持的对象 |
|:---------------|:-------------|:-----------|
| `onActivated` | 在激活工作簿时发生。 | [**Workbook**](/javascript/api/excel/excel.workbook#onActivated) |
| `onFiltered` | 当将筛选器应用于对象时发生。 | [**Table**](/javascript/api/excel/excel.table#onfiltered)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#onfiltered)、[**Worksheet**](/javascript/api/excel/excel.worksheet#onfiltered)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onfiltered) |
| `onFormulaChanged` | 更改公式时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#onFormulaChanged)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#onFormulaChanged) |

### <a name="event-triggers"></a>事件触发器

Excel 工作簿内的事件可以通过下列方式触发：

- 更改工作簿的 Excel 用户界面 (UI) 用户交互
- 更改工作簿的 Office 加载项 (JavaScript) 代码
- 更改工作簿的 VBA 加载项（宏）代码

任何符合 Excel 默认行为的更改都会在工作簿中触发一个或多个相应事件。

### <a name="lifecycle-of-an-event-handler"></a>事件处理程序的生命周期

当加载项注册事件处理程序时，将创建事件处理程序。 当加载项取消注册事件处理程序或者刷新、重新加载或关闭加载项时，将销毁事件处理程序。 事件处理程序不会暂留为 Excel 文件的一部分，也不会跨与 Excel 网页版的会话保留。

> [!CAUTION]
> 删除了注册事件的对象（例如，注册 `onChanged` 事件的表）时，事件处理程序不再触发但会保留在内存中，直到加载项或 Excel 会话刷新或关闭为止。

### <a name="events-and-coauthoring"></a>事件和共同创作

借助 [共同创作功能](co-authoring-in-excel-add-ins.md)，多个人可以共同协作，同时编辑同一个 Excel 工作簿。对于可由共同创作者触发的事件（如 `onChanged`），相应的 **Event** 对象会包含 **source** 属性，以指示事件是由当前用户在本地触发 (`event.source = Local`)，还是由远程共同创作者触发 (`event.source = Remote`)。

## <a name="register-an-event-handler"></a>注册事件处理程序

下面的代码示例为 `onChanged` 工作表中的  事件注册事件处理程序。 此代码指定 `handleChange` 函数应在工作表中的数据有变化时运行。

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

下面的代码示例为 **Sample** 工作表中的 `onSelectionChanged` 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。 它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。 请注意， `RequestContext` 用于创建事件处理程序的 需要删除它。 

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

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
