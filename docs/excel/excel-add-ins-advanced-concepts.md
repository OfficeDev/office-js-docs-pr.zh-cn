---
title: Excel JavaScript API 高级编程概念
description: 了解 Excel 加载项如何通过使用 Office JavaScript API 对象模型与 Excel 中的对象进行交互。
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: 32c46f1979b094110d32a6fcf77699eccb5d2606
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719586"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>Excel JavaScript API 高级编程概念

本文构建于 [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)中的信息之上，介绍了生成适用于 Excel 2016 或更高版本的复杂加载项所必需的一些更高级的概念。

## <a name="officejs-apis-for-excel"></a>适用于 Excel 的 Office.js API

Excel 加载项通过使用适 Office JavaScript API 与 Excel 中的对象进行交互，JavaScript API包括两个 JavaScript 对象模型：

* **Excel JavaScript API**：[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。

* **通用 API**：[通用 API](/javascript/api/office) 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。

你可能会使用 Excel JavaScript API 开发面向 Excel 2016 或更高版本的加载项中的大部分功能，同时还可以使用通用 API 中的对象。 例如：

- [Context](/javascript/api/office/office.context)：`Context` 对象表示加载项的运行时环境，并提供对 API 关键对象的访问权限。 它由工作簿配置详细信息（如 `contentLanguage` 和 `officeTheme`）组成，并提供有关加载项的运行时环境（如 `host` 和 `platform`）的信息。 此外，它还提供了 `requirements.isSetSupported()` 方法，可用于检查运行加载项的 Excel 应用程序是否支持指定的要求集。

- [Document](/javascript/api/office/office.document)：`Document` 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Excel 文件。

下图说明了可能使用 Excel JavaScript API 或公共 API 的情况。

![Excel JS API 和公共 API 之间差异的图像](../images/excel-js-api-common-api.png)

## <a name="requirement-sets"></a>要求集

要求集是指各组已命名的 API 成员。 Office 加载项可以执行运行时检查或使用清单中指定的要求集确定 Office 主机是否支持加载项所需的 API。 要确定每个受支持平台上可用的具体要求集，请参阅 [Excel JavaScript API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md)。

### <a name="checking-for-requirement-set-support-at-runtime"></a>在运行时检查要求集支持

以下代码示例显示如何确定运行加载项的主机应用程序是否支持指定的 API 要求集。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>在清单中定义要求集支持

可以在加载项清单中使用[要求元素](../reference/manifest/requirements.md)指定加载项要求激活的最小要求集和/或 API 方法。 如果 Office 主机或平台不支持清单的 `Requirements` 元素中指定的要求集或 API 方法，该加载项不会在该主机或平台中运行，而且不会显示在“我的加载项”**** 中显示的加载项列表中。

以下代码示例显示加载项清单中的 `Requirements` 元素，该元素指定应在支持 ExcelApi 要求集版本 1.3 或更高版本的所有 Office 主机应用程序中加载该加载项。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> 为了让加载项适用于 Office 主机的所有平台（如 Excel 网页版、Windows 版 Excel 和 iPad 版 Excel），建议在运行时检查是否有要求支持，而不是在清单中定义要求集支持。

### <a name="requirement-sets-for-the-officejs-common-api"></a>Office.js 通用 API 的要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。

## <a name="loading-the-properties-of-an-object"></a>加载对象的属性

在 Excel JavaScript 对象上调用 `load()` 方法指示 API 在 `sync()` 方法运行时将对象加载到 JavaScript 内存中。 `load()` 方法接受字符串（其中包含要加载的以逗号分隔的属性名称）或对象（指定要加载的属性、分页选项等）。

> [!NOTE]
> 如果对对象（或集合）调用 `load()` 方法，而未指定任何参数，将会加载对象的所有标量属性（或集合中全部对象的所有标量属性）。 为了减少 Excel 主机应用程序和加载项之间的数据传输量，应避免在没有明确指定要加载的属性的情况下调用 `load()` 方法。

### <a name="method-details"></a>方法详细信息

#### <a name="loadparam-object"></a>load(param: object)

使用参数指定的属性值和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法

```js
object.load(param);
```

#### <a name="parameters"></a>参数

|**参数**|**类型**|**说明**|
|:------------|:-------|:----------|
|`param`|object|可选。 接受属性名称作为逗号分隔的字符串或数组。 也可以传递对象来设置选择和导航属性（如下面的示例所示）。|

#### <a name="returns"></a>返回

void

#### <a name="example"></a>示例

以下代码示例通过复制另一个区域的属性来设置一个 Excel 区域的属性。 请注意，必须首先加载源对象，然后才能访问其属性值并将其写入目标区域。 此示例假定存在两个区域（**B2:E2** 和 **B7:E7**）的数据，并且这两个区域的初始格式不同。

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var sourceRange = sheet.getRange("B2:E2");
    sourceRange.load("format/fill/color, format/font/name, format/font/color");

    return ctx.sync()
        .then(function () {
            var targetRange = sheet.getRange("B7:E7");
            targetRange.set(sourceRange);
            targetRange.format.autofitColumns();

            return ctx.sync();
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="load-option-properties"></a>加载选项属性

作为调用 `load()` 方法时传递逗号分隔的字符串或数组的替代方法，可以传递一个包含以下属性的对象。

|**属性**|**类型**|**说明**|
|:-----------|:-------|:----------|
|`select`|object|包含标量属性名称的逗号分隔列表或数组。可选。|
|`expand`|object|包含导航属性名称的逗号分隔列表或数组。可选。|
|`top`|int| 指定结果中可以包含的集合项最大数量。可选。使用对象表示法选项时，仅可使用此选项。|
|`skip`|int|指定要跳过且不包含在结果中的集合中的项数目。如果指定 `top`，跳过指定数目的项目后将会启动结果集。可选。使用对象表示法选项时，仅可使用此选项。|

以下代码示例通过为集合中的每个工作表的所用区域选择 `name` 属性和 `address` 来加载工作表集合。 它还指定只能加载集合中的前五个工作表。 可以通过将 `top: 10` 和 `skip: 5` 指定为属性值来处理下一组五个工作表。

```js
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>标量和导航属性

属性分为两种类别：**标量**和**导航**。 标量属性是可分配的类型，如字符串、整数和 JSON 结构。 导航属性是只读对象和已分配字段的对象的集合，而不是直接分配属性。 例如，[Worksheet](/javascript/api/excel/excel.worksheet) 对象上的 `name`和 `position` 成员是标量属性，而 `protection` 和 `tables` 是导航属性。 [DataValidation](/javascript/api/excel/excel.datavalidation) 对象上的 `prompt` 是必须使用 JSON 对象 (`dv.prompt = { title: "MyPrompt"}`) 设置的标量属性的示例，而不是设置子属性 (`dv.prompt.title = "MyPrompt" // will not set the title`)。

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>使用 `object.load()` 的标量属性和导航属性

调用没有指定参数的 `object.load()` 方法将加载对象的所有标量属性；不会加载对象的导航属性。 此外，无法直接加载导航属性。 相反，应使用 `load()` 方法引用所需导航属性中的各个标量属性。 例如，要加载某个区域的字体名称，必须指定 `format` 和 `font` 导航属性作为 `name` 属性的路径：

```js
someRange.load("format/font/name")
```

> [!NOTE]
> 使用 Excel JavaScript API，可以通过遍历路径来设置导航属性的标量属性。 例如，可以使用 `someRange.format.font.size = 10;` 设置区域的字体大小。 在设置之前，不需要加载该属性。 

## <a name="setting-properties-of-an-object"></a>设置对象的属性

在具有嵌套导航属性的对象上设置属性可能很麻烦。 作为使用如上所述导航路径设置单个属性的替代方法，可以使用 Excel JavaScript API 中所有对象上可用的 `object.set()` 方法。 使用此方法，可以通过传递相同 Office.js 类型的另一个对象或 JavaScript 对象（其属性结构类似于调用该方法的对象的属性）一次设置对象的多个属性。

> [!NOTE]
> `set()` 方法只对主机专用 Office JavaScript API（如 Excel JavaScript API）中的对象实现。 通用（共享）API 不支持此方法。 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

对其调用此方法的对象的属性被设置为，由传入对象的相应属性所指定的值。当 `properties` 参数为 JavaScript 对象时，如果传入对象的任何属性与对其调用此方法的对象中的只读属性对应，属性会遭忽略或抛出异常，具体取决于 `options` 参数的值。

#### <a name="syntax"></a>语法

```js
object.set(properties[, options]);
```

#### <a name="parameters"></a>参数

|**参数**|**类型**|**说明**|
|:------------|:--------|:----------|
|`properties`|object|与在其上调用方法的对象相同的 Office.js 类型的对象，或属性名称及类型反映在其上调用方法的对象结构的 JavaScript 对象。|
|`options`|object|可选。只能在首个参数为 JavaScript 对象时传递。此对象可以包含下列属性：`throwOnReadOnly?: boolean`（默认值是 `true`：如果传入的 JavaScript 对象包含只读属性，将引发错误。）|

#### <a name="returns"></a>返回

void

#### <a name="example"></a>示例

下面的代码示例设置区域的多个格式属性，具体方法是调用 `set()` 方法，并传入 JavaScript 对象，其中包含可反映 `Range` 对象中属性结构的属性名称和类型。此示例假定区域 **B2:E2** 中有数据。

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods"></a>&#42;OrNullObject 方法

许多 Excel JavaScript API 方法都会在不符合 API 条件时返回异常。 例如，如果尝试通过指定工作簿中没有的工作表名称来获取工作表，`getItem()` 方法返回 `ItemNotFound` 异常。 

可以使用可用于 Excel JavaScript API 中的多种方法的 `*OrNullObject` 方法变量，而不是为此类应用场景实现复杂的异常处理逻辑。 `*OrNullObject` 方法将返回 null 对象（不是 JavaScript `null`），而不是在指定项不存在的情况下引发异常。 例如，可以在集合（如 `getItemOrNullObject()`）上调用 **** 方法，尝试从集合中检索某个项。 `getItemOrNullObject()` 方法返回指定的项（如果存在）；否则，它将返回 null 对象。 返回的 null 对象包含布尔属性 `isNullObject`，可以对其进行评估以确定该对象是否存在。

下面的代码示例尝试使用 `getItemOrNullObject()` 方法检索名为“Data”的工作表。 如果此方法返回 null 对象，需要先新建工作表，然后才能对工作表执行操作。

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
  .then(function() {
    if (dataSheet.isNullObject) {
        // Create the sheet
    }

    dataSheet.position = 1;
    //...
  })
```

## <a name="see-also"></a>另请参阅

* [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API 性能优化](performance.md)
* [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
