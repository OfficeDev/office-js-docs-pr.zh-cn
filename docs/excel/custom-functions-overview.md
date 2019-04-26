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
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="4c638-103">在 Excel 中创建自定义函数（预览）</span><span class="sxs-lookup"><span data-stu-id="4c638-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="4c638-104">开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="4c638-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="4c638-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="4c638-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="4c638-106">本文介绍了如何在 Excel 中创建自定义函数。</span><span class="sxs-lookup"><span data-stu-id="4c638-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="4c638-107">下图演示最终用户将自定义函数插入到 Excel 工作表单元格的过程。</span><span class="sxs-lookup"><span data-stu-id="4c638-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="4c638-108">`CONTOSO.ADD42` 自定义函数旨在向用户指定作为函数输入参数的数字对添加 42。</span><span class="sxs-lookup"><span data-stu-id="4c638-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="4c638-109">以下代码定义 `ADD42` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="4c638-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="4c638-110">本文后面的[已知问题](#known-issues)部分指定自定义函数的当前限制。</span><span class="sxs-lookup"><span data-stu-id="4c638-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="4c638-111">自定义函数加载项项目的组件</span><span class="sxs-lookup"><span data-stu-id="4c638-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="4c638-112">如果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，会发现它可创建全面控制函数、任务窗格和加载项的文件。</span><span class="sxs-lookup"><span data-stu-id="4c638-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="4c638-113">我们将专注于对自定义函数至关重要的文件：</span><span class="sxs-lookup"><span data-stu-id="4c638-113">We'll concentrate on the files that are important to custom functions:</span></span> 

| <span data-ttu-id="4c638-114">文件</span><span class="sxs-lookup"><span data-stu-id="4c638-114">File</span></span> | <span data-ttu-id="4c638-115">文件格式</span><span class="sxs-lookup"><span data-stu-id="4c638-115">File format</span></span> | <span data-ttu-id="4c638-116">说明</span><span class="sxs-lookup"><span data-stu-id="4c638-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="4c638-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="4c638-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="4c638-118">或</span><span class="sxs-lookup"><span data-stu-id="4c638-118">or</span></span><br/><span data-ttu-id="4c638-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="4c638-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="4c638-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="4c638-120">JavaScript</span></span><br/><span data-ttu-id="4c638-121">或</span><span class="sxs-lookup"><span data-stu-id="4c638-121">or</span></span><br/><span data-ttu-id="4c638-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="4c638-122">TypeScript</span></span> | <span data-ttu-id="4c638-123">包含定义自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="4c638-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="4c638-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="4c638-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="4c638-125">HTML</span><span class="sxs-lookup"><span data-stu-id="4c638-125">HTML</span></span> | <span data-ttu-id="4c638-126">提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。</span><span class="sxs-lookup"><span data-stu-id="4c638-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="4c638-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="4c638-127">**./manifest.xml**</span></span> | <span data-ttu-id="4c638-128">XML</span><span class="sxs-lookup"><span data-stu-id="4c638-128">XML</span></span> | <span data-ttu-id="4c638-129">指定加载项中所有自定义函数的命名空间以及此表中前面列出的 JavaScript 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="4c638-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="4c638-130">它还列出了加载项可能使用的其他文件的位置，如任务窗格文件和命令文件。</span><span class="sxs-lookup"><span data-stu-id="4c638-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="4c638-131">脚本文件</span><span class="sxs-lookup"><span data-stu-id="4c638-131">Script file</span></span>

<span data-ttu-id="4c638-132">脚本文件（Yo Office 生成器创建的项目中的 **./src/functions/functions.js** 或 **./src/functions/functions.ts**）包含定义自定义函数的代码、定义函数的注释，并将自定义函数名称关联到 JSON 元数据文件中的对象。</span><span class="sxs-lookup"><span data-stu-id="4c638-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions, comments which define the function, and associates the names of the custom functions to objects in the JSON metadata file.</span></span>

<span data-ttu-id="4c638-133">以下代码定义自定义函数 `add`，然后指定该函数的关联信息。</span><span class="sxs-lookup"><span data-stu-id="4c638-133">The following code defines the custom function `add`  and then specifies association information for the function.</span></span> <span data-ttu-id="4c638-134">有关关联函数的详细信息，请参阅[自定义函数最佳做法](custom-functions-best-practices.md#associating-function-names-with-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="4c638-134">For more information on associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

<span data-ttu-id="4c638-135">下面的代码还提供了定义函数的代码注释。</span><span class="sxs-lookup"><span data-stu-id="4c638-135">The following code also provides code comments which define the function.</span></span> <span data-ttu-id="4c638-136">首先声明所需的 `@customfunction` 注释，指示这是一个自定义函数。</span><span class="sxs-lookup"><span data-stu-id="4c638-136">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="4c638-137">此外，你将注意到声明了两个参数，即 `first` 和 `second`，后跟其 `description` 属性。</span><span class="sxs-lookup"><span data-stu-id="4c638-137">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="4c638-138">最后提供了 `returns` 描述。</span><span class="sxs-lookup"><span data-stu-id="4c638-138">Finally, a `returns` description is given.</span></span> <span data-ttu-id="4c638-139">有关自定义函数所需注释的更多信息，请参阅[为自定义函数生成 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="4c638-139">For more information about what comments are required for your custom function, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

### <a name="manifest-file"></a><span data-ttu-id="4c638-140">清单文件</span><span class="sxs-lookup"><span data-stu-id="4c638-140">Manifest file</span></span>

<span data-ttu-id="4c638-141">定义自定义函数的加载项的 XML 清单文件（Yo Office 生成器创建的项目中的 **./manifest.xml**）指定加载项中所有自定义函数的命名空间以及 JavaScript、JSON 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="4c638-141">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="4c638-142">下面的基本 XML 标记显示了 `<ExtensionPoint>` 和 `<Resources>` 元素的一个示例，必须在加载项清单中包含这些元素才能启用自定义函数。</span><span class="sxs-lookup"><span data-stu-id="4c638-142">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="4c638-143">如果使用 Yo Office 生成器，生成的自定义函数文件将包含更复杂的清单文件，可以在[此 Github 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml)中对其进行比较。</span><span class="sxs-lookup"><span data-stu-id="4c638-143">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="4c638-144">在自定义函数 JavaScript、JSON 和 HTML 文件的清单文件中指定的 URL 必须可公开访问，并具有相同的子域。</span><span class="sxs-lookup"><span data-stu-id="4c638-144">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

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
> <span data-ttu-id="4c638-145">Excel 中的函数在前面追加 XML 清单文件中指定的命名空间作为前缀。</span><span class="sxs-lookup"><span data-stu-id="4c638-145">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="4c638-146">函数的命名空间在函数名称之前，并用句点分隔。</span><span class="sxs-lookup"><span data-stu-id="4c638-146">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="4c638-147">例如，若要在 Excel 工作表的单元格中调用函数 `ADD42`，需输入 `=CONTOSO.ADD42`，因为 `CONTOSO` 是命名空间，`ADD42` 是 JSON 文件中指定的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="4c638-147">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="4c638-148">命名空间旨在作为公司或加载项的标识符使用。</span><span class="sxs-lookup"><span data-stu-id="4c638-148">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="4c638-149">命名空间只能包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="4c638-149">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="4c638-150">声明可变函数</span><span class="sxs-lookup"><span data-stu-id="4c638-150">Declaring a volatile function</span></span>

<span data-ttu-id="4c638-151">[可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)是指其值时刻更改的函数（即使此函数的自变量均未更改）。</span><span class="sxs-lookup"><span data-stu-id="4c638-151">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="4c638-152">每当 Excel 重新计算时，这些函数即会重新计算。</span><span class="sxs-lookup"><span data-stu-id="4c638-152">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="4c638-153">例如，假设某个单元格调用函数 `NOW`。</span><span class="sxs-lookup"><span data-stu-id="4c638-153">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="4c638-154">每当调用 `NOW` 时，它将自动返回当前的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4c638-154">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="4c638-155">Excel 包含多个内置可变函数，例如 `RAND` 和 `TODAY`。</span><span class="sxs-lookup"><span data-stu-id="4c638-155">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="4c638-156">有关 Excel 可变函数的完整列表，请参阅[可变函数和非可变函数](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions)。</span><span class="sxs-lookup"><span data-stu-id="4c638-156">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="4c638-157">借助自定义函数，可以创建自己的可变函数。处理日期、时间、随机数字和建模时，可能会使用可变函数。</span><span class="sxs-lookup"><span data-stu-id="4c638-157">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="4c638-158">例如，Monte Carlo 模拟需要生成随机输入，来确定最佳解决方案。</span><span class="sxs-lookup"><span data-stu-id="4c638-158">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="4c638-159">若要声明可变函数，则在 JSON 元数据文件内相应函数的 `options` 对象中添加 `"volatile": true`，如下面的代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="4c638-159">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="4c638-160">请注意，无法同时将一个函数标记为 `"streaming": true` 和 `"volatile": true`；当同时将这两者标记为 `true` 时，将忽略可变选项。</span><span class="sxs-lookup"><span data-stu-id="4c638-160">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

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

## <a name="saving-and-sharing-state"></a><span data-ttu-id="4c638-161">保存和共享状态</span><span class="sxs-lookup"><span data-stu-id="4c638-161">Saving and sharing state</span></span>

<span data-ttu-id="4c638-162">自定义函数可以将数据保存在全局 JavaScript 变量中，可用于后续调用。</span><span class="sxs-lookup"><span data-stu-id="4c638-162">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="4c638-163">当用户从多个单元格调用同一个自定义函数时，保存状态非常有用，因为函数的所有实例都可以访问该状态。</span><span class="sxs-lookup"><span data-stu-id="4c638-163">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="4c638-164">例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。</span><span class="sxs-lookup"><span data-stu-id="4c638-164">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="4c638-165">下面的代码示例演示温度流式处理函数的实现过程，该函数在全局范围内保存状态。</span><span class="sxs-lookup"><span data-stu-id="4c638-165">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="4c638-166">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="4c638-166">Note the following about this code:</span></span>

- <span data-ttu-id="4c638-167">`streamTemperature` 函数每秒更新单元格中显示的温度值，并使用 `savedTemperatures` 变量作为其数据源。</span><span class="sxs-lookup"><span data-stu-id="4c638-167">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="4c638-168">因为 `streamTemperature` 是一个流式处理函数，它将实现一个取消处理程序，当函数被取消时该处理程序将运行。</span><span class="sxs-lookup"><span data-stu-id="4c638-168">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="4c638-169">如果用户从 Excel 中的多个单元格调用 `streamTemperature` 函数，则 `streamTemperature` 函数在每次运行时都会从相同的 `savedTemperatures` 变量读取数据。</span><span class="sxs-lookup"><span data-stu-id="4c638-169">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="4c638-170">`refreshTemperature` 函数每秒读取特定温度计的温度，并将结果存储在 `savedTemperatures` 变量中。</span><span class="sxs-lookup"><span data-stu-id="4c638-170">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="4c638-171">因为 `refreshTemperature` 函数不在 Excel 中向最终用户显示，所以不需要在 JSON 文件中注册。</span><span class="sxs-lookup"><span data-stu-id="4c638-171">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="coauthoring"></a><span data-ttu-id="4c638-172">共同创作</span><span class="sxs-lookup"><span data-stu-id="4c638-172">Coauthoring</span></span>

<span data-ttu-id="4c638-173">借助 Excel Online 和 Excel for Windows 以及 Office 365 订阅，可以共同创作文档，此功能可与自定义函数结合使用。</span><span class="sxs-lookup"><span data-stu-id="4c638-173">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="4c638-174">如果你的工作簿使用自定义函数，系统会提示你的同事加载自定义函数的加载项。</span><span class="sxs-lookup"><span data-stu-id="4c638-174">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="4c638-175">当你们均加载此加载项后，自定义函数会通过共同创作共享结果。</span><span class="sxs-lookup"><span data-stu-id="4c638-175">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="4c638-176">若要详细了解共同创作，请参阅[关于 Excel 中的共同创作](/office/vba/excel/concepts/about-coauthoring-in-excel)。</span><span class="sxs-lookup"><span data-stu-id="4c638-176">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="4c638-177">使用数据区域</span><span class="sxs-lookup"><span data-stu-id="4c638-177">Working with ranges of data</span></span>

<span data-ttu-id="4c638-178">自定义函数可以接受数据区域作为输入参数，也可以返回数据区域。</span><span class="sxs-lookup"><span data-stu-id="4c638-178">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="4c638-179">在 JavaScript，数据区域表示为一个二维数组。</span><span class="sxs-lookup"><span data-stu-id="4c638-179">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="4c638-180">例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。</span><span class="sxs-lookup"><span data-stu-id="4c638-180">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="4c638-181">下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。</span><span class="sxs-lookup"><span data-stu-id="4c638-181">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="4c638-182">请注意，在此函数的 JSON 元数据中，将参数的 `type` 属性设置为 `matrix`。</span><span class="sxs-lookup"><span data-stu-id="4c638-182">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="4c638-183">确定调用自定义函数的单元格</span><span class="sxs-lookup"><span data-stu-id="4c638-183">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="4c638-184">在某些情况下，需要获取调用自定义函数的单元格地址。</span><span class="sxs-lookup"><span data-stu-id="4c638-184">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="4c638-185">这在以下类型的应用场景中非常有用：</span><span class="sxs-lookup"><span data-stu-id="4c638-185">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="4c638-186">设置区域格式：将单元格地址用作键，以便将信息存储到 [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data) 中。</span><span class="sxs-lookup"><span data-stu-id="4c638-186">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="4c638-187">然后，使用 Excel 中的 [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) 从 `AsyncStorage` 加载该键。</span><span class="sxs-lookup"><span data-stu-id="4c638-187">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="4c638-188">显示缓存值：如果脱机使用函数，将显示 `AsyncStorage` 中使用 `onCalculated` 存储的缓存值。</span><span class="sxs-lookup"><span data-stu-id="4c638-188">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="4c638-189">协调：使用单元格地址发现原始单元格，以帮助你在处理时进行协调。</span><span class="sxs-lookup"><span data-stu-id="4c638-189">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="4c638-190">仅当函数 JSON 元数据文件中的 `requiresAddress` 被标记为 `true` 时，才会公开与单元格地址相关的信息。</span><span class="sxs-lookup"><span data-stu-id="4c638-190">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="4c638-191">以下示例诠释了此情况：</span><span class="sxs-lookup"><span data-stu-id="4c638-191">The following sample gives an example of this:</span></span>

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

<span data-ttu-id="4c638-192">此外，需要在脚本文件（**./src/functions/functions.js** 或 **./src/functions/functions.ts**）中添加 `getAddress` 函数，以查找单元格地址。</span><span class="sxs-lookup"><span data-stu-id="4c638-192">In the script file (**./src/functions/functions.js** or **./src/functions/functions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="4c638-193">此函数可能会使用参数，如以下示例 `parameter1` 所示。</span><span class="sxs-lookup"><span data-stu-id="4c638-193">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="4c638-194">最后一个参数始终为 `invocationContext`，该对象包含 JSON 元数据文件中的 `requiresAddress` 被标记为 `true` 时 Excel 传递的单元格位置。</span><span class="sxs-lookup"><span data-stu-id="4c638-194">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="4c638-195">默认情况下，从 `getAddress` 函数返回的值遵循以下格式：`SheetName!CellNumber`。</span><span class="sxs-lookup"><span data-stu-id="4c638-195">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="4c638-196">例如，如果名为“Expense”的工作表中的 B2 单元格调用了函数，则返回的值为 `Expenses!B2`。</span><span class="sxs-lookup"><span data-stu-id="4c638-196">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="4c638-197">已知问题</span><span class="sxs-lookup"><span data-stu-id="4c638-197">Known issues</span></span>

<span data-ttu-id="4c638-198">在 [Excel 自定义功能 GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/issues)上查看已知问题。</span><span class="sxs-lookup"><span data-stu-id="4c638-198">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="4c638-199">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4c638-199">See also</span></span>

* [<span data-ttu-id="4c638-200">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="4c638-200">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="4c638-201">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="4c638-201">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="4c638-202">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="4c638-202">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="4c638-203">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="4c638-203">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="4c638-204">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="4c638-204">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="4c638-205">自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="4c638-205">Custom functions debugging</span></span>](custom-functions-debugging.md)
