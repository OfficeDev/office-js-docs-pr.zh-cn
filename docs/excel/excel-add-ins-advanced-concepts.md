---
title: 使用 Excel JavaScript API 的高级编程概念
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 09f2d95e4cf7631b519f00cddee265dbf697e07e
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505886"
---
# <a name="advanced-programming-concepts-with-the-excel-javascript-api"></a>使用 Excel JavaScript API 的高级编程概念

本文根据 [Excel JavaScript API 核心方案](excel-add-ins-core-concepts.md)中的信息，介绍了生成适用于 Excel 2016 的复杂加载项所必需的部分高级方案。

## <a name="officejs-apis-for-excel"></a>适用于 Excel 的 Office.js API

Excel 加载项通过使用适用于 Office 的 JavaScript API 与 Excel 中的对象进行交互，该 API 包括两个 JavaScript 对象模型：

* **Excel JavaScript API**：自 Office 2016 引入的 [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js) 提供了强类型的对象，可用于访问工作表、区域、表格、图表等。 

* **公用 API**：自 Office 2013 引入的公用 API（也称为[共享 API](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)）可用于访问诸如 Word、Excel 和 PowerPoint 等多种宿主应用程序通用的 UI、对话框和客户端设置等功能。

尽管大概率会使用 Excel JavaScript API 开发面向 Excel 2016 或后续版本的加载项的大部分功能，还将使用到共享 API 中的对象。 例如：

- [Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js)： **Context** 对象表示加载项的运行时环境，并提供对 API 的关键对象的访问。它包含工作簿配置详细信息如 `contentLanguage` 和 `officeTheme`，还提供了有关加载项运行时环境的信息，如 `host` 和 `platform`。此外，它提供了 `requirements.isSetSupported()` 方法，您可用于检查运行加载项的 Excel 应用程序是否支持指定的要求集。 

- [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js)：**Document** 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Excel 文件。 

## <a name="requirement-sets"></a>要求集

要求集是具名的 API 成员的组合。Office 加载项可以执行运行时检查或使用清单中指定的要求集，以确定 Office 宿主是否支持加载项所需的 API。若要了解每个受支持的平台的可用特定要求集，请参阅 [Excel 的 JavaScript API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js)。

### <a name="checking-for-requirement-set-support-at-runtime"></a>在运行时检查要求集支持

以下代码示例显示如何确定运行加载项的宿主应用程序是否支持指定的 API 要求集。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>在清单中定义要求集支持

你可以在加载项清单中使用 [Requirements 元素](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements?view=office-js)指定最低要求集和/或加载项需要激活的 API 方法。如果 Office 宿主或平台不支持清单的 **Requiremsnts** 元素中指定的要求集或 API 方法, 加载项不会在宿主或平台中运行，并不会显示于**我的加载项**列表中。 

以下代码示例显示加载项清单中的 **Requirements** 元素，该元素指定应在支持 ExcelApi 要求集版本 1.3 或更高版本的所有 Office 宿主应用程序中加载该加载项。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> 若要让加载项适用于 Office 宿主的所有平台（如 Excel for Windows、Excel Online 和 Excel for iPad），建议在运行时检查要求是否支持，而不是在清单中定义要求集支持。

### <a name="requirement-sets-for-the-officejs-common-api"></a>Office.js 公用 API 的要求集

有关公用 API 要求集的信息，请参阅 [Office 公用 API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)。

## <a name="loading-the-properties-of-an-object"></a>加载对象的属性

调用 Excel JavaScript 对象上的 `load()` 方法指示 API 在  `sync()` 方法运行时将对象加载到 JavaScript 内存中时。 `load()` 方法接受一个包含逗号分隔的加载属性的名称字符串，或者指定要加载属性和分页选项等的对象。 

> [!NOTE]
> 如果您调用对象（或集合）的 `load()` 方法时未指定任何参数，将加载对象的所有标量属性 （或集合中的所有对象的所有标量属性） 。为了减少的 Excel 主机应用程序和加载项之间的数据传输，应避免呼叫 `load()` 方法时未明确指定要加载的属性。

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
|`param`|object|可选。接受以逗号分隔字符串或数组指定的参数和关系名称。此外，可以通过传递对象来设置选择和导航属性 （如下面的示例中所示）。|

#### <a name="returns"></a>返回

void

#### <a name="example"></a>示例

下面的代码示例通过复制其他区域的属性来设置一个 Excel 区域的属性。请注意，必须首先加载的源对象，才可以访问其属性值并写入到目标区域。本示例假定的两个范围 （**b2: e2** 和 **B7:E7**）均有数据且最初设置了不同的格式。

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

作为调用 `load()` 方法时传递逗号分隔字符串或数组的替代方法，可以传递一个包含以下属性的对象。 

|**属性**|**类型**|**说明**|
|:-----------|:-------|:----------|
|`select`|object|包含参数/关系名称的逗号分隔列表或数组。可选。|
|`expand`|object|包含关系名称的逗号分隔列表或数组。可选。|
|`top`|int| 指定结果中可以包含的集合项最大数量。可选。使用对象表示法选项时，仅可使用此选项。|
|`skip`|int|指定要跳过且不包含在结果中的集合中的项数目。如果指定 `top`，跳过指定数目的项目后将会启动结果集。可选。使用对象表示法选项时，仅可使用此选项。|

下面的代码示例通过选择 `name` 属性和集合中每个工作表所使用范围的 `address` 加载工作表集合。它还指定只应加载集合中的前五个工作表。可以通过将  `top: 10` 和 `skip: 5`  指定为属性值来处理下一组五个工作表。 

```js 
myWorksheets.load({
    select: 'name, userRange/address',
    expand: 'tables',
    top: 5,
    skip: 0
});
```

## <a name="scalar-and-navigation-properties"></a>标量和导航属性 

在 Excel 的 JavaScript API 参考文档中，您可能会发现对象成员分为两个类别： **属性**和**关系**。对象的属性是字符串、一个整数或布尔值等标量，对象的关系 （也称为导航属性）是作为对象或对象集合的成员。例如，[Worksheet](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) 对象的 `name` 和 `position` 成员是标量属性，而 `protection` 和 `tables` 是关系（导航属性）。 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>标量属性和导航属性与 `object.load()`

调用没有指定参数的 `object.load()` 方法将加载对象的所有标量属性；不会加载对象的导航属性。此外，不能直接加载导航属性。相反，应使用 `load()` 方法，以引用所需的导航属性中的各个标量属性。例如，若要加载范围的字体名称，则必须作为 **name** 属性的路径而指定 **format** 和 **font** 导航属性：

```js
someRange.load("format/font/name")
```

> [!NOTE]
> 使用 Excel 的 JavaScript API，你可以通过遍历路径设置导航属性的标量属性。例如，可以使用 `someRange.format.font.size = 10;` 设置范围的字体大小。不需要在设置之前加载属性。 

## <a name="setting-properties-of-an-object"></a>设置对象的属性

在具有嵌套导航属性的对象上设置属性可能很麻烦。作为使用如上所述导航路径设置单个属性的替代方法，可以使用 Excel JavaScript API 中所有对象上可用的  `object.set()` 方法，可对 Excel 的 JavaScript API 中的所有对象。使用此方法时，可以通过传递相同 Office.js 类型的另一个对象或 JavaScript 对象（其属性结构类似于调用该方法的对象的属性）一次设置对象的多个属性。

> [!NOTE]
> `set()` 方法只对宿主应用程序 Office JavaScript API（如 Excel JavaScript API）中的对象实现。通用（共享） API 不支持此方法。 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

对其调用此方法的对象的属性被设置为由传入对象的相应属性所指定的值。当 `properties` 参数为 JavaScript 对象时，如果传入对象的任何属性与对其调用此方法的对象中的只读属性对应，属性会遭忽略或抛出异常，具体取决于 `options` 参数的值。

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

下面的代码示例设置范围的多个格式属性，具体方法是调用 `set()` 方法，并传入 JavaScript 对象，其中包含可反映 **Range** 对象中属性结构的属性名称和类型。此示例假定范围 **B2:E2** 中有数据。

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
## <a name="42ornullobject-methods"></a>*OrNullObject 方法

不满足 API 的条件时，许多 Excel 的 JavaScript API 方法将返回异常。例如，如果尝试通过指定工作簿中不存在的工作表名称中获取工作表，`getItem()` 方法将返回 `ItemNotFound` 异常。 

可以使用可用于 Excel JavaScript API 中的多种方法的 `*OrNullObject` 方法变体，而不是为此类应用场景实现复杂的异常处理逻辑。在指定的项不存在时，`*OrNullObject` 方法将返回 null 对象（非 JavaScript `null`）而不是引发异常。例如，可以对 **Worksheets** 等集合调用 `getItemOrNullObject()` 方法，以尝试从集合中检索项目。 指定项如果存在，`getItemOrNullObject()` 方法将返回该项，否则它将返回 null 对象。返回的 null 对象包含布尔属性 `isNullObject`，用于确定对象是否存在。

下面的代码示例尝试使用 `getItemOrNullObject()` 方法检索名为“Data”的工作表。如果该方法返回 null 对象，需要先新建工作表，然后才能对工作表执行操作。

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
 
* [使用 Excel JavaScript API 的基本编程概念](excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API 性能优化](performance.md)
* [Excel JavaScript API 参考](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
