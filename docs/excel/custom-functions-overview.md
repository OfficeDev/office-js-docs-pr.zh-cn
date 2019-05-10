---
ms.date: 05/03/2019
description: 在 Excel 中使用 JavaScript 创建自定义函数。
title: 在 Excel 中创建自定义函数
localization_priority: Priority
ms.openlocfilehash: 5a31cc8ddfe98b880ab09803c7c0b7b615ba85db
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659647"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="77dfc-103">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="77dfc-103">Create custom functions in Excel</span></span> 

<span data-ttu-id="77dfc-104">开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="77dfc-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="77dfc-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="77dfc-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="77dfc-106">本文介绍了如何在 Excel 中创建自定义函数。</span><span class="sxs-lookup"><span data-stu-id="77dfc-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="77dfc-107">以下动态图像显示调用你使用 JavaScript 或 Typescript 创建的函数的工作簿。</span><span class="sxs-lookup"><span data-stu-id="77dfc-107">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="77dfc-108">在此示例中，自定义函数 `=MYFUNCTION.SPHEREVOLUME` 计算球的体积。</span><span class="sxs-lookup"><span data-stu-id="77dfc-108">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolume.gif" />

<span data-ttu-id="77dfc-109">以下代码定义 `=MYFUNCTION.SPHEREVOLUME` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="77dfc-109">The following code defines the `=MYFUNCTION.SPHEREVOLUME` custom function.</span></span>

```js
/**
 * Returns the volume of a sphere. 
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
CustomFunctions.associate("SPHEREVOLUME", sphereVolume)
```

> [!NOTE]
> <span data-ttu-id="77dfc-110">本文后面的[已知问题](#known-issues)部分指定自定义函数的当前限制。</span><span class="sxs-lookup"><span data-stu-id="77dfc-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="77dfc-111">如何在代码中定义自定义函数</span><span class="sxs-lookup"><span data-stu-id="77dfc-111">How a custom function is defined in code</span></span>

<span data-ttu-id="77dfc-112">如果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，会发现它可创建全面控制函数、任务窗格和加载项的文件。</span><span class="sxs-lookup"><span data-stu-id="77dfc-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll find that it creates files which control your functions, your task pane, and your add-in overall.</span></span> <span data-ttu-id="77dfc-113">我们将专注于对自定义函数至关重要的文件：</span><span class="sxs-lookup"><span data-stu-id="77dfc-113">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="77dfc-114">文件</span><span class="sxs-lookup"><span data-stu-id="77dfc-114">File</span></span> | <span data-ttu-id="77dfc-115">文件格式</span><span class="sxs-lookup"><span data-stu-id="77dfc-115">File format</span></span> | <span data-ttu-id="77dfc-116">说明</span><span class="sxs-lookup"><span data-stu-id="77dfc-116">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="77dfc-117">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="77dfc-117">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="77dfc-118">或</span><span class="sxs-lookup"><span data-stu-id="77dfc-118">or</span></span><br/><span data-ttu-id="77dfc-119">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="77dfc-119">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="77dfc-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="77dfc-120">JavaScript</span></span><br/><span data-ttu-id="77dfc-121">或</span><span class="sxs-lookup"><span data-stu-id="77dfc-121">or</span></span><br/><span data-ttu-id="77dfc-122">TypeScript</span><span class="sxs-lookup"><span data-stu-id="77dfc-122">TypeScript</span></span> | <span data-ttu-id="77dfc-123">包含定义自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="77dfc-123">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="77dfc-124">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="77dfc-124">**./src/functions/functions.html**</span></span> | <span data-ttu-id="77dfc-125">HTML</span><span class="sxs-lookup"><span data-stu-id="77dfc-125">HTML</span></span> | <span data-ttu-id="77dfc-126">提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。</span><span class="sxs-lookup"><span data-stu-id="77dfc-126">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="77dfc-127">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="77dfc-127">**./manifest.xml**</span></span> | <span data-ttu-id="77dfc-128">XML</span><span class="sxs-lookup"><span data-stu-id="77dfc-128">XML</span></span> | <span data-ttu-id="77dfc-129">指定加载项中所有自定义函数的命名空间以及此表中前面列出的 JavaScript 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="77dfc-129">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript and HTML files that are listed previously in this table.</span></span> <span data-ttu-id="77dfc-130">它还列出了加载项可能使用的其他文件的位置，如任务窗格文件和命令文件。</span><span class="sxs-lookup"><span data-stu-id="77dfc-130">It also lists the locations of other files your add-in might make use of, such as the task pane files and command files.</span></span> |

### <a name="script-file"></a><span data-ttu-id="77dfc-131">脚本文件</span><span class="sxs-lookup"><span data-stu-id="77dfc-131">Script file</span></span>

<span data-ttu-id="77dfc-132">脚本文件（**./src/functions/functions.js** 或 **./src/functions/functions.ts**）包含定义自定义函数的代码、定义函数的注释，并将自定义函数名称关联到 JSON 元数据文件中的对象。</span><span class="sxs-lookup"><span data-stu-id="77dfc-132">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions, comments which define the function, and associates the names of the custom functions to objects in the JSON metadata file.</span></span>

<span data-ttu-id="77dfc-133">以下代码定义 `add` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="77dfc-133">The following code defines the `add` custom function.</span></span> <span data-ttu-id="77dfc-134">代码注释用于生成描述 Excel 自定义函数的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="77dfc-134">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="77dfc-135">首先声明所需的 `@customfunction` 注释，指示这是一个自定义函数。</span><span class="sxs-lookup"><span data-stu-id="77dfc-135">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="77dfc-136">此外，你将注意到声明了两个参数，即 `first` 和 `second`，后跟其 `description` 属性。</span><span class="sxs-lookup"><span data-stu-id="77dfc-136">Additionally, you'll notice two parameters are declared, `first` and `second`, which are followed by their `description` properties.</span></span> <span data-ttu-id="77dfc-137">最后提供了 `returns` 描述。</span><span class="sxs-lookup"><span data-stu-id="77dfc-137">Finally, a `returns` description is given.</span></span> <span data-ttu-id="77dfc-138">要详细了解自定义函数需要哪些注释，请参阅[为自定义函数创建 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="77dfc-138">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="77dfc-139">以下代码还调用 `CustomFunctions.associate("ADD", add)`，以将函数 `add()` 与其在 JSON 元数据文件 `ADD` 中的 ID 关联。</span><span class="sxs-lookup"><span data-stu-id="77dfc-139">The following code also calls `CustomFunctions.associate("ADD", add)` to associate the function `add()` with its ID in the JSON metadata file `ADD`.</span></span> <span data-ttu-id="77dfc-140">有关关联函数的详细信息，请参阅[自定义函数最佳做法](custom-functions-best-practices.md#associating-function-names-with-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="77dfc-140">For more information on associating functions, see [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata).</span></span>

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="77dfc-141">请注意，控制自定义函数运行时加载的 **functions.html** 文件必须链接至自定义函数的当前 CDN。</span><span class="sxs-lookup"><span data-stu-id="77dfc-141">Note that the **functions.html** file, which governs the loading of the custom functions runtime, must link to the current CDN for custom functions.</span></span> <span data-ttu-id="77dfc-142">准备有当前版本的 Yo Office 生成器的项目引用正确的 CDN。</span><span class="sxs-lookup"><span data-stu-id="77dfc-142">Projects prepared with the current version of the Yo Office generator reference the correct CDN.</span></span> <span data-ttu-id="77dfc-143">如果更新 2019 年 3 月或更早的自定义函数项目，则需要将以下代码复制到 **functions.html** 页面。</span><span class="sxs-lookup"><span data-stu-id="77dfc-143">If you are retrofitting a previous custom function project from March 2019 or earlier, you need to copy in the code below to the **functions.html** page.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a><span data-ttu-id="77dfc-144">清单文件</span><span class="sxs-lookup"><span data-stu-id="77dfc-144">Manifest file</span></span>

<span data-ttu-id="77dfc-145">定义自定义函数的加载项的 XML 清单文件（Yo Office 生成器创建的项目中的 **./manifest.xml**）指定加载项中所有自定义函数的命名空间以及 JavaScript、JSON 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="77dfc-145">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> 

<span data-ttu-id="77dfc-146">下面的基本 XML 标记显示了 `<ExtensionPoint>` 和 `<Resources>` 元素的一个示例，必须在加载项清单中包含这些元素才能启用自定义函数。</span><span class="sxs-lookup"><span data-stu-id="77dfc-146">The following basic XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span> <span data-ttu-id="77dfc-147">如果使用 Yo Office 生成器，生成的自定义函数文件将包含更复杂的清单文件，可以在[此 Github 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml)中对其进行比较。</span><span class="sxs-lookup"><span data-stu-id="77dfc-147">If using the Yo Office generator, your generated custom function files will contain a more complex manifest file, which you can compare on [this Github repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/generate-metadata/manifest.xml).</span></span>

> [!NOTE] 
> <span data-ttu-id="77dfc-148">在自定义函数 JavaScript、JSON 和 HTML 文件的清单文件中指定的 URL 必须可公开访问，并具有相同的子域。</span><span class="sxs-lookup"><span data-stu-id="77dfc-148">The URLs specified in the manifest file for the custom functions JavaScript, JSON, and HTML files must be publicly accessible and have the same subdomain.</span></span>

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
> <span data-ttu-id="77dfc-149">Excel 中的函数在前面追加 XML 清单文件中指定的命名空间作为前缀。</span><span class="sxs-lookup"><span data-stu-id="77dfc-149">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="77dfc-150">函数的命名空间在函数名称之前，并用句点分隔。</span><span class="sxs-lookup"><span data-stu-id="77dfc-150">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="77dfc-151">例如，若要在 Excel 工作表的单元格中调用函数 `ADD42`，需输入 `=CONTOSO.ADD42`，因为 `CONTOSO` 是命名空间，`ADD42` 是 JSON 文件中指定的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="77dfc-151">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="77dfc-152">命名空间旨在作为公司或加载项的标识符使用。</span><span class="sxs-lookup"><span data-stu-id="77dfc-152">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="77dfc-153">命名空间只能包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="77dfc-153">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="77dfc-154">共同创作</span><span class="sxs-lookup"><span data-stu-id="77dfc-154">Coauthoring</span></span>

<span data-ttu-id="77dfc-155">借助 Excel Online 和 Excel for Windows 以及 Office 365 订阅，可以共同创作文档，此功能可与自定义函数结合使用。</span><span class="sxs-lookup"><span data-stu-id="77dfc-155">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="77dfc-156">如果你的工作簿使用自定义函数，系统会提示你的同事加载自定义函数的加载项。</span><span class="sxs-lookup"><span data-stu-id="77dfc-156">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="77dfc-157">当你们均加载此加载项后，自定义函数会通过共同创作共享结果。</span><span class="sxs-lookup"><span data-stu-id="77dfc-157">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="77dfc-158">若要详细了解共同创作，请参阅[关于 Excel 中的共同创作](/office/vba/excel/concepts/about-coauthoring-in-excel)。</span><span class="sxs-lookup"><span data-stu-id="77dfc-158">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="77dfc-159">已知问题</span><span class="sxs-lookup"><span data-stu-id="77dfc-159">Known issues</span></span>

<span data-ttu-id="77dfc-160">在 [Excel 自定义功能 GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/issues)上查看已知问题。</span><span class="sxs-lookup"><span data-stu-id="77dfc-160">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="77dfc-161">后续步骤</span><span class="sxs-lookup"><span data-stu-id="77dfc-161">Next steps</span></span>

<span data-ttu-id="77dfc-162">想要试用自定义函数？</span><span class="sxs-lookup"><span data-stu-id="77dfc-162">Want to try out custom functions?</span></span> <span data-ttu-id="77dfc-163">检查简单的[自定义函数入门](../quickstarts/excel-custom-functions-quickstart.md)或更深入的[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)（如果还没有）。</span><span class="sxs-lookup"><span data-stu-id="77dfc-163">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span> 

<span data-ttu-id="77dfc-164">另一个尝试自定义函数的简单方法就是使用[脚本实验室](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)，这是一个允许您在 Excel 中试验自定义函数的加载项。</span><span class="sxs-lookup"><span data-stu-id="77dfc-164">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="77dfc-165">可以尝试创建自己的自定义函数或使用提供的示例。</span><span class="sxs-lookup"><span data-stu-id="77dfc-165">You can try out creating your own custom function or play with the provided samples.</span></span>

<span data-ttu-id="77dfc-166">准备详细了解自定义函数的功能？</span><span class="sxs-lookup"><span data-stu-id="77dfc-166">Ready to read more about the capabilities custom functions?</span></span> <span data-ttu-id="77dfc-167">了解[自定义函数架构](custom-functions-architecture.md)的概述。</span><span class="sxs-lookup"><span data-stu-id="77dfc-167">Learn about an overview of [the custom functions architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="77dfc-168">另请参阅</span><span class="sxs-lookup"><span data-stu-id="77dfc-168">See also</span></span> 
* [<span data-ttu-id="77dfc-169">自定义函数要求</span><span class="sxs-lookup"><span data-stu-id="77dfc-169">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="77dfc-170">命名准则</span><span class="sxs-lookup"><span data-stu-id="77dfc-170">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="77dfc-171">最佳做法</span><span class="sxs-lookup"><span data-stu-id="77dfc-171">Best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="77dfc-172">让自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="77dfc-172">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
