---
ms.date: 05/17/2020
description: 为 Office 加载项创建 Excel 自定义函数
title: 在 Excel 中创建自定义函数
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 42ace6208abbd95d0f538345a1f5b5cc15ba1823
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093460"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="50d74-103">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="50d74-103">Create custom functions in Excel</span></span>

<span data-ttu-id="50d74-104">开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="50d74-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="50d74-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="50d74-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="50d74-106">以下动态图像显示调用你使用 JavaScript 或 Typescript 创建的函数的工作簿。</span><span class="sxs-lookup"><span data-stu-id="50d74-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="50d74-107">在此示例中，自定义函数 `=MYFUNCTION.SPHEREVOLUME` 计算球的体积。</span><span class="sxs-lookup"><span data-stu-id="50d74-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="50d74-108">以下代码定义 `=MYFUNCTION.SPHEREVOLUME` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="50d74-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

```js
/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!NOTE]
> <span data-ttu-id="50d74-109">本文后面的[已知问题](#known-issues)部分指定自定义函数的当前限制。</span><span class="sxs-lookup"><span data-stu-id="50d74-109">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="50d74-110">如何在代码中定义自定义函数</span><span class="sxs-lookup"><span data-stu-id="50d74-110">How a custom function is defined in code</span></span>

<span data-ttu-id="50d74-111">如果使用[Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，它将创建用于控制函数和任务窗格的文件。</span><span class="sxs-lookup"><span data-stu-id="50d74-111">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="50d74-112">我们将专注于对自定义函数至关重要的文件：</span><span class="sxs-lookup"><span data-stu-id="50d74-112">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="50d74-113">文件</span><span class="sxs-lookup"><span data-stu-id="50d74-113">File</span></span> | <span data-ttu-id="50d74-114">文件格式</span><span class="sxs-lookup"><span data-stu-id="50d74-114">File format</span></span> | <span data-ttu-id="50d74-115">说明</span><span class="sxs-lookup"><span data-stu-id="50d74-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="50d74-116">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="50d74-116">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="50d74-117">或</span><span class="sxs-lookup"><span data-stu-id="50d74-117">or</span></span><br/><span data-ttu-id="50d74-118">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="50d74-118">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="50d74-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="50d74-119">JavaScript</span></span><br/><span data-ttu-id="50d74-120">或</span><span class="sxs-lookup"><span data-stu-id="50d74-120">or</span></span><br/><span data-ttu-id="50d74-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="50d74-121">TypeScript</span></span> | <span data-ttu-id="50d74-122">包含定义自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="50d74-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="50d74-123">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="50d74-123">**./src/functions/functions.html**</span></span> | <span data-ttu-id="50d74-124">HTML</span><span class="sxs-lookup"><span data-stu-id="50d74-124">HTML</span></span> | <span data-ttu-id="50d74-125">提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。</span><span class="sxs-lookup"><span data-stu-id="50d74-125">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="50d74-126">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="50d74-126">**./manifest.xml**</span></span> | <span data-ttu-id="50d74-127">XML</span><span class="sxs-lookup"><span data-stu-id="50d74-127">XML</span></span> | <span data-ttu-id="50d74-128">指定自定义函数使用的多个文件的位置，例如自定义函数 JavaScript、JSON 和 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="50d74-128">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="50d74-129">此外，它还列出了任务窗格文件、命令文件的位置，并指定了自定义函数应使用的运行时。</span><span class="sxs-lookup"><span data-stu-id="50d74-129">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="50d74-130">脚本文件</span><span class="sxs-lookup"><span data-stu-id="50d74-130">Script file</span></span>

<span data-ttu-id="50d74-131">脚本文件 (**./src/functions/functions.js** or **./src/functions/functions.ts**) 包含定义自定义函数的代码以及定义函数的注释。</span><span class="sxs-lookup"><span data-stu-id="50d74-131">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="50d74-132">以下代码定义 `add` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="50d74-132">The following code defines the custom function `add`.</span></span> <span data-ttu-id="50d74-133">代码注释用于生成描述 Excel 自定义函数的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="50d74-133">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="50d74-134">首先声明所需的 `@customfunction` 注释，指示这是一个自定义函数。</span><span class="sxs-lookup"><span data-stu-id="50d74-134">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="50d74-135">接下来，先声明两个参数， `first` 然后再 `second` 键入它们的 `description` 属性。</span><span class="sxs-lookup"><span data-stu-id="50d74-135">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="50d74-136">最后提供了 `returns` 描述。</span><span class="sxs-lookup"><span data-stu-id="50d74-136">Finally, a `returns` description is given.</span></span> <span data-ttu-id="50d74-137">要详细了解自定义函数需要哪些注释，请参阅[为自定义函数创建 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="50d74-137">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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
```

### <a name="manifest-file"></a><span data-ttu-id="50d74-138">清单文件</span><span class="sxs-lookup"><span data-stu-id="50d74-138">Manifest file</span></span>

<span data-ttu-id="50d74-139">用于定义自定义函数的加载项的 XML 清单文件，该项目是在 Yo Office 生成器创建的项目中 (**。/manifest.xml。**) 执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="50d74-139">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="50d74-140">定义自定义函数的命名空间。</span><span class="sxs-lookup"><span data-stu-id="50d74-140">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="50d74-141">命名空间将自己添加到自定义函数中，以帮助客户将您的函数标识为外接程序的一部分。</span><span class="sxs-lookup"><span data-stu-id="50d74-141">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="50d74-142">使用 `<ExtensionPoint>` 和 `<Resources>` 元素对于自定义函数清单而言是唯一的。</span><span class="sxs-lookup"><span data-stu-id="50d74-142">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="50d74-143">这些元素包含有关 JavaScript、JSON 和 HTML 文件的位置的信息。</span><span class="sxs-lookup"><span data-stu-id="50d74-143">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="50d74-144">指定要用于自定义函数的运行时。</span><span class="sxs-lookup"><span data-stu-id="50d74-144">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="50d74-145">我们建议始终使用共享运行时，除非您有其他运行时的特定需求，因为共享运行时允许在函数和任务窗格之间共享数据。</span><span class="sxs-lookup"><span data-stu-id="50d74-145">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span>

<span data-ttu-id="50d74-146">如果使用 Yo Office 生成器创建文件，建议将清单调整为使用共享运行时，因为这不是这些文件的默认值。</span><span class="sxs-lookup"><span data-stu-id="50d74-146">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="50d74-147">若要更改清单，请按照[配置 Excel 外接程序中的说明使用共享的 JavaScript 运行时](./configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="50d74-147">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](./configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="50d74-148">若要查看示例加载项中的完整工作清单，请参阅[此 Github 存储库](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml)。</span><span class="sxs-lookup"><span data-stu-id="50d74-148">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="50d74-149">共同创作</span><span class="sxs-lookup"><span data-stu-id="50d74-149">Coauthoring</span></span>

<span data-ttu-id="50d74-150">Web 上的 Excel 和连接到 Microsoft 365 订阅的 Windows 允许您在 Excel 中 coauthor。</span><span class="sxs-lookup"><span data-stu-id="50d74-150">Excel on the web and Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="50d74-151">如果您的工作簿使用自定义函数，则将提示您的合著同事加载自定义函数的外接程序。</span><span class="sxs-lookup"><span data-stu-id="50d74-151">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="50d74-152">一旦您加载了加载项，自定义函数将通过共同创作来共享结果。</span><span class="sxs-lookup"><span data-stu-id="50d74-152">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="50d74-153">若要详细了解共同创作，请参阅[关于 Excel 中的共同创作](/office/vba/excel/concepts/about-coauthoring-in-excel)。</span><span class="sxs-lookup"><span data-stu-id="50d74-153">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="50d74-154">已知问题</span><span class="sxs-lookup"><span data-stu-id="50d74-154">Known issues</span></span>

<span data-ttu-id="50d74-155">在 [Excel 自定义功能 GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/issues)上查看已知问题。</span><span class="sxs-lookup"><span data-stu-id="50d74-155">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="50d74-156">后续步骤</span><span class="sxs-lookup"><span data-stu-id="50d74-156">Next steps</span></span>

<span data-ttu-id="50d74-157">想要试用自定义函数？</span><span class="sxs-lookup"><span data-stu-id="50d74-157">Want to try out custom functions?</span></span> <span data-ttu-id="50d74-158">检查简单的[自定义函数入门](../quickstarts/excel-custom-functions-quickstart.md)或更深入的[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)（如果还没有）。</span><span class="sxs-lookup"><span data-stu-id="50d74-158">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="50d74-159">另一个尝试自定义函数的简单方法就是使用[脚本实验室](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)，这是一个允许您在 Excel 中试验自定义函数的加载项。</span><span class="sxs-lookup"><span data-stu-id="50d74-159">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="50d74-160">可以尝试创建自己的自定义函数或使用提供的示例。</span><span class="sxs-lookup"><span data-stu-id="50d74-160">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="50d74-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="50d74-161">See also</span></span> 
* [<span data-ttu-id="50d74-162">自定义函数要求</span><span class="sxs-lookup"><span data-stu-id="50d74-162">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="50d74-163">命名准则</span><span class="sxs-lookup"><span data-stu-id="50d74-163">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="50d74-164">让自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="50d74-164">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
