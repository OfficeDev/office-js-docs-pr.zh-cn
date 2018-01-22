# <a name="excel-javascript-api-advanced-concepts"></a>Excel JavaScript API 高级概念

本文根据 [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)中的信息，介绍了生成适用于 Excel 2016 的复杂加载项所必需的一些更高级的概念。 

## <a name="officejs-apis-for-excel"></a>适用于 Excel 的 Office.js API

Excel 加载项通过使用适用于 Office 的 JavaScript API 与 Excel 中的对象进行交互，适用于 Office 的 JavaScript API包括两个 JavaScript 对象模型：

* **Excel JavaScript API**：[Excel JavaScript API](http://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) 随 Office 2016 一起引入，提供了强类型的对象，可用于访问工作表、区域、表格、图表等。 

* **公用 API**：公用 API（也称为[共享 API](http://dev.office.com/reference/add-ins/javascript-api-for-office)）随 Office 2013 一起引入，可用于访问诸如 Word、Excel 和 PowerPoint 等多种主机应用程序通用的 UI、对话框和客户端设置等功能。

你可能会使用 Excel JavaScript API 开发面向 Excel 2016 的加载项中的大部分功能，同时还可以使用共享 API 中的对象。 例如：

- [Context](http://dev.office.com/reference/add-ins/shared/context)：**Context** 对象表示加载项的运行时环境，并提供对 API 的关键对象的访问。 它由工作簿配置详细信息（如 `contentLanguage` 和 `officeTheme`）组成，并提供有关加载项的运行时环境（如 `host` 和 `platform`）的信息。 此外，它还提供了 `requirements.isSetSupported()` 方法，可用于检查运行加载项的 Excel 应用程序是否支持指定的要求集。 

- [Document](http://dev.office.com/reference/add-ins/shared/document)：**Document** 对象提供 `getFileAsync()` 方法，用于下载运行加载项的 Excel 文件。 

## <a name="requirement-sets"></a>要求集

要求集是指各组已命名的 API 成员。 Office 加载项可以执行运行时检查或使用清单中指定的要求集确定 Office 主机是否支持加载项所需的 API。 要确定每个受支持平台上可用的具体要求集，请参阅 [Excel JavaScript API 要求集](http://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)。

### <a name="checking-for-requirement-set-support-at-runtime"></a>在运行时检查要求集支持

以下代码示例显示如何确定运行加载项的主机应用程序是否支持指定的 API 要求集。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>在清单中定义要求集支持

可以在加载项清单中使用[要求元素](http://dev.office.com/reference/add-ins/manifest/requirements)指定加载项要求激活的最小要求集和/或 API 方法。 如果 Office 主机或平台不支持清单的 **Requirements** 元素中指定的要求集或 API 方法，该加载项不会在该主机或平台中运行，而且不会显示在“我的加载项”****中显示的加载项列表中。 

以下代码示例显示加载项清单中的 **Requirements** 元素，该元素指定应在支持 ExcelApi 要求集版本 1.3 或更高版本的所有 Office 主机应用程序中加载该加载项。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> **注意**：要使加载项可用于 Office 主机的所有平台，如 Excel for Windows、Excel Online 和 Excel for iPad，建议在运行时检查要求支持，而不是在清单中定义要求集支持。

### <a name="requirement-sets-for-the-officejs-common-api"></a>Office.js 公用 API 的要求集

有关公用 API 要求集的信息，请参阅 [Office 公用 API 要求集](http://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)。

## <a name="loading-the-properties-of-an-object"></a>加载对象的属性

在 Excel JavaScript 对象上调用 `load()` 方法指示 API 在 `sync()` 方法运行时将对象加载到 JavaScript 内存中。 `load()` 方法接受字符串（其中包含要加载的以逗号分隔的属性名称）或对象（指定要加载的属性、分页选项等）。 

> **注意**：如果在不指定任何参数的情况下在对象（或集合）上调用 `load()` 方法，则会加载对象的所有标量属性（或集合中所有对象的所有标量属性）。 为了减少 Excel 主机应用程序和加载项之间的数据传输量，应避免在没有明确指定要加载的属性的情况下调用 `load()` 方法。

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
|`param`|object|可选。 接受参数和关系名称作为逗号分隔的字符串或数组。 也可以传递对象来设置选择和导航属性（如下面的示例所示）。|

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
|`select`|object|包含逗号分隔的列表或参数/关系名称的数组。 可选。|
|`expand`|object|包含逗号分隔的列表或关系名称的数组。 可选。|
|`top`|int| 指定结果中可以包含的集合项最大数量。可选。使用对象表示法选项时，仅可使用此选项。|
|`skip`|int|指定要跳过且不包含在结果中的集合中的项数目。 如果指定 `top`，跳过指定数目的项目后将会启动结果集。 可选。 只有在使用对象表示法选项时，才能使用此选项。|

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

在 Excel JavaScript API 参考文档中，你可能会注意到，对象成员分为两类：**属性**和**关系**。 对象的属性是一个标量成员（如字符串、整数或布尔值），而对象的关系（也称为“导航属性”）是一个对象/对象集合成员。 例如，[Worksheet](http://dev.office.com/reference/add-ins/excel/worksheet) 对象中的 `name` 和 `position` 成员是标量属性，而 `protection` 和 `tables` 是关系（导航属性）。 

### <a name="scalar-properties-and-navigation-properties-with-objectload"></a>使用 `object.load()` 的标量属性和导航属性

调用没有指定参数的 `object.load()` 方法将加载对象的所有标量属性；不会加载对象的导航属性。 此外，无法直接加载导航属性。 相反，应使用 `load()` 方法引用所需导航属性中的各个标量属性。 例如，要加载某个区域的字体名称，必须指定 **format** 和 **font** 导航属性作为 **name** 属性的路径：

```js
someRange.load("format/font/name")
```

> **注意**：使用 Excel JavaScript API，可以通过遍历路径来设置导航属性的标量属性。 例如，可以使用 `someRange.format.font.size = 10;` 设置区域的字体大小。 在设置之前，不需要加载该属性。 

## <a name="setting-properties-of-an-object"></a>设置对象的属性

在具有嵌套导航属性的对象上设置属性可能很麻烦。 作为使用如上所述导航路径设置单个属性的替代方法，可以使用 Excel JavaScript API 中所有对象上可用的 `object.set()` 方法。 使用此方法，可以通过传递相同 Office.js 类型的另一个对象或 JavaScript 对象（其属性结构类似于调用该方法的对象的属性）一次设置对象的多个属性。

> **注意**：`set()` 方法仅适用于特定于主机的 Office JavaScript API（如 Excel JavaScript API）中的对象。 公用（共享）API 不支持此方法。 

### <a name="set-properties-object-options-object"></a>set (properties: object, options: object)

将在其上调用方法的对象的属性设置为由传入对象的相应属性指定的值。 如果 `properties` 参数是一个 JavaScript 对象，则传入对象的任何与在其上调用方法的对象中的只读属性相对应的属性将被忽略或导致抛出异常，具体取决于 `options` 参数的值。

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

以下代码示例通过调用 `set()` 方法并传入具有可反映 **Range** 对象中属性结构的属性名称和类型的 JavaScript 对象来设置区域的多个格式属性。 此示例假定区域 **B2:E2** 中包含数据。

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

如果不符合 API 的条件，许多 Excel JavaScript API 方法将返回异常。 例如，如果你尝试通过指定工作簿中不存在的工作表名称来获取工作表，则 `getItem()` 方法将返回 `ItemNotFound` 异常。 

可以使用可用于 Excel JavaScript API 中的多种方法的 `*OrNullObject` 方法变量，而不是为此类应用场景实现复杂的异常处理逻辑。 `*OrNullObject` 方法将返回 null 对象（不是 JavaScript `null`），而不是在指定项不存在的情况下引发异常。 例如，可以在集合（如 **Worksheets**）上调用 `getItemOrNullObject()` 方法，尝试从集合中检索某个项。 `getItemOrNullObject()` 方法返回指定的项（如果存在）；否则，它将返回 null 对象。 返回的 null 对象包含布尔属性 `isNullObject`，可以对其进行评估以确定该对象是否存在。

下面的代码示例尝试使用 `getItemOrNullObject()` 方法检索名为“Data”的工作表。 如果该方法返回 null 对象，则需要创建新工作表之后，才能在该工作表上执行操作。

```js
let dataSheet = context.workbook.worksheets.getItemOrNullObject("Data"); 
if (dataSheet.isNullObject) { 
    // Create the sheet
}

dataSheet.position = 1;
//...
```

## <a name="additional-resources"></a>其他资源
 
* [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API 参考](http://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
