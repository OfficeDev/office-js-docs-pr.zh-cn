---
ms.date: 03/29/2019
description: 在 Excel 中使用 JavaScript 创建自定义函数。
title: 在 Excel 中创建自定义函数（预览）
localization_priority: Priority
ms.openlocfilehash: 7a461728061ace532a11a8473d27ec4340eebb97
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448469"
---
# <a name="create-custom-functions-in-excel-preview"></a>在 Excel 中创建自定义函数（预览）

开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。 Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。 本文介绍了如何在 Excel 中创建自定义函数。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

下图演示最终用户将自定义函数插入到 Excel 工作表单元格的过程。 `CONTOSO.ADD42` 自定义函数旨在向用户指定作为函数输入参数的数字对添加 42。

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

以下代码定义 `ADD42` 自定义函数。

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> 本文后面的[已知问题](#known-issues)部分指定自定义函数的当前限制。

## <a name="components-of-a-custom-functions-add-in-project"></a>自定义函数加载项项目的组件

如果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，会发现它可创建全面控制函数、任务窗格和加载项的文件。 我们将专注于对自定义函数至关重要的文件： 

| 文件 | 文件格式 | 说明 |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>或<br/>**./src/functions/functions.ts** | JavaScript<br/>或<br/>TypeScript | 包含定义自定义函数的代码。 |
| **./src/functions/functions.html** | HTML | 提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。 |
| **./manifest.xml** | XML | 指定加载项中所有自定义函数的命名空间以及此表中前面列出的 JavaScript 和 HTML 文件的位置。 它还列出了加载项可能使用的其他文件的位置，如任务窗格文件和命令文件。 |

### <a name="script-file"></a>脚本文件

脚本文件（Yo Office 生成器创建的项目中的 **./src/functions/functions.js** 或 **./src/functions/functions.ts**）包含定义自定义函数的代码、定义函数的注释，并将自定义函数名称关联到 JSON 元数据文件中的对象。

以下代码定义自定义函数 `add`，然后指定该函数的关联信息。 有关关联函数的详细信息，请参阅[自定义函数最佳做法](custom-functions-best-practices.md#associating-function-names-with-json-metadata)。

下面的代码还提供了定义函数的代码注释。 首先声明所需的 `@customfunction` 注释，指示这是一个自定义函数。 此外，你将注意到声明了两个参数，即 `first` 和 `second`，后跟其 `description` 属性。 最后提供了 `returns` 描述。 有关自定义函数所需注释的更多信息，请参阅[为自定义函数生成 JSON 元数据](custom-functions-json-autogeneration.md)。

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

### <a name="manifest-file"></a>清单文件

定义自定义函数的加载项的 XML 清单文件（Yo Office 生成器创建的项目中的 **./manifest.xml**）指定加载项中所有自定义函数的命名空间以及 JavaScript、JSON 和 HTML 文件的位置。 

下面的基本 XML 标记显示了 `<ExtensionPoint>` 和 `<Resources>` 元素的一个示例，必须在加载项清单中包含这些元素才能启用自定义函数。 如果使用 Yo Office 生成器，生成的自定义函数文件将包含更复杂的清单文件，可以在[此 Github 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml)中对其进行比较。

> [!NOTE] 
> 在自定义函数 JavaScript、JSON 和 HTML 文件的清单文件中指定的 URL 必须可公开访问，并具有相同的子域。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> Excel 中的函数在前面追加 XML 清单文件中指定的命名空间作为前缀。 函数的命名空间在函数名称之前，并用句点分隔。 例如，若要在 Excel 工作表的单元格中调用函数 `ADD42`，需输入 `=CONTOSO.ADD42`，因为 `CONTOSO` 是命名空间，`ADD42` 是 JSON 文件中指定的函数的名称。 命名空间旨在作为公司或加载项的标识符使用。 命名空间只能包含字母数字字符和句点。

## <a name="declaring-a-volatile-function"></a>声明可变函数

[可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)是指其值时刻更改的函数（即使此函数的自变量均未更改）。 每当 Excel 重新计算时，这些函数即会重新计算。 例如，假设某个单元格调用函数 `NOW`。 每当调用 `NOW` 时，它将自动返回当前的日期和时间。

Excel 包含多个内置可变函数，例如 `RAND` 和 `TODAY`。 有关 Excel 可变函数的完整列表，请参阅[可变函数和非可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)。

借助自定义函数，可以创建自己的可变函数。处理日期、时间、随机数字和建模时，可能会使用可变函数。 例如，Monte Carlo 模拟需要生成随机输入，来确定最佳解决方案。

若要声明可变函数，则在 JSON 元数据文件内相应函数的 `options` 对象中添加 `"volatile": true`，如下面的代码示例所示。 请注意，无法同时将一个函数标记为 `"streaming": true` 和 `"volatile": true`；当同时将这两者标记为 `true` 时，将忽略可变选项。

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## <a name="saving-and-sharing-state"></a>保存和共享状态

自定义函数可以将数据保存在全局 JavaScript 变量中，可用于后续调用。 当用户从多个单元格调用同一个自定义函数时，保存状态非常有用，因为函数的所有实例都可以访问该状态。 例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。

下面的代码示例演示温度流式处理函数的实现过程，该函数在全局范围内保存状态。 关于此代码，请注意以下几点：

- `streamTemperature` 函数每秒更新单元格中显示的温度值，并使用 `savedTemperatures` 变量作为其数据源。

- 因为 `streamTemperature` 是一个流式处理函数，它将实现一个取消处理程序，当函数被取消时该处理程序将运行。

- 如果用户从 Excel 中的多个单元格调用 `streamTemperature` 函数，则 `streamTemperature` 函数在每次运行时都会从相同的 `savedTemperatures` 变量读取数据。 

- `refreshTemperature` 函数每秒读取特定温度计的温度，并将结果存储在 `savedTemperatures` 变量中。 因为 `refreshTemperature` 函数不在 Excel 中向最终用户显示，所以不需要在 JSON 文件中注册。

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="coauthoring"></a>共同创作

借助 Excel Online 和 Excel for Windows 以及 Office 365 订阅，可以共同创作文档，此功能可与自定义函数结合使用。 如果你的工作簿使用自定义函数，系统会提示你的同事加载自定义函数的加载项。 当你们均加载此加载项后，自定义函数会通过共同创作共享结果。

若要详细了解共同创作，请参阅[关于 Excel 中的共同创作](/office/vba/excel/concepts/about-coauthoring-in-excel)。

## <a name="working-with-ranges-of-data"></a>使用数据区域

自定义函数可以接受数据区域作为输入参数，也可以返回数据区域。 在 JavaScript，数据区域表示为一个二维数组。

例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。 下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。 请注意，在此函数的 JSON 元数据中，将参数的 `type` 属性设置为 `matrix`。

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="determine-which-cell-invoked-your-custom-function"></a>确定调用自定义函数的单元格

在某些情况下，需要获取调用自定义函数的单元格地址。 这在以下类型的应用场景中非常有用：

- 设置区域格式：将单元格地址用作键，以便将信息存储到 [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data) 中。 然后，使用 Excel 中的 [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) 从 `AsyncStorage` 加载该键。
- 显示缓存值：如果脱机使用函数，将显示 `AsyncStorage` 中使用 `onCalculated` 存储的缓存值。
- 协调：使用单元格地址发现原始单元格，以帮助你在处理时进行协调。

仅当函数 JSON 元数据文件中的 `requiresAddress` 被标记为 `true` 时，才会公开与单元格地址相关的信息。 以下示例诠释了此情况：

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

此外，需要在脚本文件（**./src/functions/functions.js** 或 **./src/functions/functions.ts**）中添加 `getAddress` 函数，以查找单元格地址。 此函数可能会使用参数，如以下示例 `parameter1` 所示。 最后一个参数始终为 `invocationContext`，该对象包含 JSON 元数据文件中的 `requiresAddress` 被标记为 `true` 时 Excel 传递的单元格位置。

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

默认情况下，从 `getAddress` 函数返回的值遵循以下格式：`SheetName!CellNumber`。 例如，如果名为“Expense”的工作表中的 B2 单元格调用了函数，则返回的值为 `Expenses!B2`。

## <a name="known-issues"></a>已知问题

在 [Excel 自定义功能 GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/issues)上查看已知问题。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
* [自定义函数调试](custom-functions-debugging.md)
