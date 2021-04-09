---
title: 使用 Excel JavaScript API 处理日期
description: 将Moment-MSDate Excel JavaScript API 的插件用于日期。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d3f59e5daad042541bd933fb4e644d40f27a6e5e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652802"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a><span data-ttu-id="242f5-103">使用 Excel JavaScript API 和 Moment-MSDate插件处理日期</span><span class="sxs-lookup"><span data-stu-id="242f5-103">Work with dates using the Excel JavaScript API and the Moment-MSDate plug-in</span></span>

<span data-ttu-id="242f5-104">本文提供的代码示例显示如何使用 Excel JavaScript API 和 [Moment-MSDate 插件处理日期](https://www.npmjs.com/package/moment-msdate)。</span><span class="sxs-lookup"><span data-stu-id="242f5-104">This article provides code samples that show how to work with dates using the Excel JavaScript API and the [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate).</span></span> <span data-ttu-id="242f5-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="242f5-105">For the complete list of properties and methods that the `Range` object supports, see the [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a><span data-ttu-id="242f5-106">使用Moment-MSDate插件处理日期</span><span class="sxs-lookup"><span data-stu-id="242f5-106">Use the Moment-MSDate plug-in to work with dates</span></span>

<span data-ttu-id="242f5-107">[时刻 JavaScript 库](https://momentjs.com/)提供了使用日期和时间戳的便捷方式。</span><span class="sxs-lookup"><span data-stu-id="242f5-107">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="242f5-108">[Moment-MSDate 插件](https://www.npmjs.com/package/moment-msdate)可将时刻格式转换为 Excel 所需的格式。</span><span class="sxs-lookup"><span data-stu-id="242f5-108">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="242f5-109">这是 [NOW 函数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)返回的相同格式。</span><span class="sxs-lookup"><span data-stu-id="242f5-109">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="242f5-110">以下代码演示如何将 **B4** 区域设置为时刻的时间戳。</span><span class="sxs-lookup"><span data-stu-id="242f5-110">The following code shows how to set the range at **B4** to a moment's timestamp.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="242f5-111">下面的代码示例演示了一种类似的技术，用于从单元格重新获取日期并将其转换为 `Moment` 或其他格式。</span><span class="sxs-lookup"><span data-stu-id="242f5-111">The following code sample demonstrates a similar technique to get the date back out of the cell and convert it to a `Moment` or other format.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="242f5-112">外接程序必须设置范围的格式，以更可读的形式显示日期。</span><span class="sxs-lookup"><span data-stu-id="242f5-112">Your add-in has to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="242f5-113">例如，显示 `"[$-409]m/d/yy h:mm AM/PM;@"` "12/3/18 3：57 PM"。</span><span class="sxs-lookup"><span data-stu-id="242f5-113">For example, `"[$-409]m/d/yy h:mm AM/PM;@"` displays "12/3/18 3:57 PM".</span></span> <span data-ttu-id="242f5-114">有关日期和时间数字格式详细信息，请参阅查看自定义数字格式指南文章中的"日期和时间 [格式指南](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) "。</span><span class="sxs-lookup"><span data-stu-id="242f5-114">For more information about date and time number formats, see "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>


## <a name="see-also"></a><span data-ttu-id="242f5-115">另请参阅</span><span class="sxs-lookup"><span data-stu-id="242f5-115">See also</span></span>

- [<span data-ttu-id="242f5-116">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="242f5-116">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="242f5-117">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="242f5-117">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="242f5-118"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="242f5-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
