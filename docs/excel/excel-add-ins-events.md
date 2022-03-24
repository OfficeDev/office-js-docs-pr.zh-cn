---
title: 使用 Excel JavaScript API 处理事件
description: JavaScript 对象Excel列表。 这包括有关使用事件处理程序和相关模式的信息。
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: c15beba846fc5348143b63dfb07321b6dad01ea2
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745012"
---
# <a name="work-with-events-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理事件

本文介绍了与处理 Excel 中事件相关的重要概念，并提供了代码示例，以展示如何使用 Excel JavaScript API 注册事件处理程序、处理事件和删除事件处理程序。

## <a name="events-in-excel"></a>Excel 中的事件

每当 Excel 工作簿中出现某种类型的更改时，就会触发事件通知。 使用 Excel JavaScript API，可以注册事件处理程序，以便加载项能够在发生特定事件时自动运行指定的函数。 下列事件暂不受支持。

| 事件 | 说明 | 支持的对象 |
|:---------------|:-------------|:-----------|
| `onActivated` | 激活对象时发生。 | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member)、[**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)、[**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member) |
| `onActivated` | 在激活工作簿时发生。 | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member) |
| `onAdded` | 当向集合中添加对象时发生。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member)、 [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member)、 [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member)、 [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member) |
| `onAutoSaveSettingChanged` | 在工作簿上更改 `autoSave` 设置时发生。 | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member) |
| `onCalculated` | 工作表完成计算（或集合的所有工作表都已完成）时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member) |
| `onChanged` | 在单个单元格或批注的数据发生更改时发生。 | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member)、 [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member)、 [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member)、 [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member)、 [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member) |
| `onColumnSorted` | 在已对一个或多个列进行排序时发生。 这是从左到右排序操作的结果。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member) |
| `onDataChanged` | 当绑定内的数据或格式变化时发生。 | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-ondatachanged-member) |
| `onDeactivated` | 停用对象时发生。 | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member)、[**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member)、[**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)、[**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member) |
| `onDeleted` | 当从集合中删除对象时发生。 | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member)、 [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member)、 [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member)、 [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member) |
| `onFormatChanged` | 在工作表上的格式变化时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member) |
| `onFormulaChanged` | 更改公式时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member) |
| `onProtectionChanged` | 工作表保护状态更改时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member) |
| `onRowHiddenChanged` | 在特定工作表上的行隐藏状态更改时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member) |
| `onRowSorted` | 在已对一个或多个行进行排序时发生。 这是从上到下排序操作的结果。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member) |
| `onSelectionChanged` | 当活动单元格或选定范围更改时发生。 | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-onselectionchanged-member)、[**Table**](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)、[**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onselectionchanged-member)、[**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)[**、WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member) |
| `onSettingsChanged` | 当文档中的设置变化时发生。 | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member) |
| `onSingleClicked` | 在工作表中进行左键单击/点击操作时发生。 | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member) |

### <a name="events-in-preview"></a>预览版中的事件

> [!NOTE]
> 以下事件当前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| 事件 | 说明 | 支持的对象 |
|:---------------|:-------------|:-----------|
| `onFiltered` | 当将筛选器应用于对象时发生。 | [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member)、[**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member)、[**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member)、[**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member) |

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
await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    await context.sync();
    console.log("Event handler successfully registered for onChanged event in the worksheet.");
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>处理事件

如上一示例所示，注册事件处理程序时，指定函数应在指定事件发生时运行。 可以将函数设计为执行方案所需的任何操作。 下面的代码示例展示了事件处理程序函数如何直接将事件信息写入控制台。

```js
async function handleChange(event) {
    await Excel.run(async (context) => {
        await context.sync();        
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);       
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>删除事件处理程序

下面的代码示例为 **Sample** 工作表中的 `onSelectionChanged` 事件注册事件处理程序，并将 `handleSelectionChange` 函数定义为在事件发生时运行。 它还定义了随后可以调用的 `remove()` 函数，以删除相应事件处理程序。 请注意， `RequestContext` 用于创建事件处理程序的 需要删除它。

```js
let eventResult;

async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    await context.sync();
    console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
  });
}

async function handleSelectionChange(event) {
  await Excel.run(async (context) => {
    await context.sync();
    console.log("Address of current selection: " + event.address);
  });
}

async function remove() {
  await Excel.run(eventResult.context, async (context) => {
    eventResult.remove();
    await context.sync();
    
    eventResult = null;
    console.log("Event handler successfully removed.");
  });
}
```

## <a name="enable-and-disable-events"></a>启用和禁用事件

可以通过禁用事件来改进加载项性能。
例如，你的应用可能永远不需要接收事件，也可能在执行多个实体的批量编辑时忽略事件。

启用和禁用事件是在[运行时](/javascript/api/excel/excel.runtime)级别进行的。
`enableEvents` 属性确定是否触发事件并激活其处理程序。

以下代码示例展示了如何打开和关闭事件。

```js
await Excel.run(async (context) => {
    context.runtime.load("enableEvents");
    await context.sync();

    let eventBoolean = !context.runtime.enableEvents;
    context.runtime.enableEvents = eventBoolean;
    if (eventBoolean) {
        console.log("Events are currently on.");
    } else {
        console.log("Events are currently off.");
    }
    
    await context.sync();
});
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
