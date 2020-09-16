---
ms.date: 08/13/2020
description: 为 Office 加载项创建 Excel 自定义函数。
title: 在 Excel 中创建自定义函数
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 731e8d99a36cfef7d125838c67efcdd7a77b4bb1
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819558"
---
# <a name="create-custom-functions-in-excel"></a><span data-ttu-id="e50e6-103">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="e50e6-103">Create custom functions in Excel</span></span>

<span data-ttu-id="e50e6-104">开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="e50e6-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="e50e6-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="e50e6-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="e50e6-106">以下动态图像显示调用你使用 JavaScript 或 Typescript 创建的函数的工作簿。</span><span class="sxs-lookup"><span data-stu-id="e50e6-106">The following animated image shows your workbook calling a function you've created with JavaScript or Typescript.</span></span> <span data-ttu-id="e50e6-107">在此示例中，自定义函数 `=MYFUNCTION.SPHEREVOLUME` 计算球的体积。</span><span class="sxs-lookup"><span data-stu-id="e50e6-107">In this example, the custom function `=MYFUNCTION.SPHEREVOLUME` calculates the volume of a sphere.</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="e50e6-108">以下代码定义 `=MYFUNCTION.SPHEREVOLUME` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="e50e6-108">The following code defines the custom function `=MYFUNCTION.SPHEREVOLUME`.</span></span>

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
> <span data-ttu-id="e50e6-109">本文后面的[已知问题](#known-issues)部分指定自定义函数的当前限制。</span><span class="sxs-lookup"><span data-stu-id="e50e6-109">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="how-a-custom-function-is-defined-in-code"></a><span data-ttu-id="e50e6-110">如何在代码中定义自定义函数</span><span class="sxs-lookup"><span data-stu-id="e50e6-110">How a custom function is defined in code</span></span>

<span data-ttu-id="e50e6-111">如果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，则它可创建控制你的函数和任务窗格的文件。</span><span class="sxs-lookup"><span data-stu-id="e50e6-111">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, it creates files which control your functions and task pane.</span></span> <span data-ttu-id="e50e6-112">我们将专注于对自定义函数至关重要的文件：</span><span class="sxs-lookup"><span data-stu-id="e50e6-112">We'll concentrate on the files that are important to custom functions:</span></span>

| <span data-ttu-id="e50e6-113">文件</span><span class="sxs-lookup"><span data-stu-id="e50e6-113">File</span></span> | <span data-ttu-id="e50e6-114">文件格式</span><span class="sxs-lookup"><span data-stu-id="e50e6-114">File format</span></span> | <span data-ttu-id="e50e6-115">说明</span><span class="sxs-lookup"><span data-stu-id="e50e6-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="e50e6-116">**./src/functions/functions.js**</span><span class="sxs-lookup"><span data-stu-id="e50e6-116">**./src/functions/functions.js**</span></span><br/><span data-ttu-id="e50e6-117">或</span><span class="sxs-lookup"><span data-stu-id="e50e6-117">or</span></span><br/><span data-ttu-id="e50e6-118">**./src/functions/functions.ts**</span><span class="sxs-lookup"><span data-stu-id="e50e6-118">**./src/functions/functions.ts**</span></span> | <span data-ttu-id="e50e6-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="e50e6-119">JavaScript</span></span><br/><span data-ttu-id="e50e6-120">或</span><span class="sxs-lookup"><span data-stu-id="e50e6-120">or</span></span><br/><span data-ttu-id="e50e6-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="e50e6-121">TypeScript</span></span> | <span data-ttu-id="e50e6-122">包含定义自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="e50e6-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="e50e6-123">**./src/functions/functions.html**</span><span class="sxs-lookup"><span data-stu-id="e50e6-123">**./src/functions/functions.html**</span></span> | <span data-ttu-id="e50e6-124">HTML</span><span class="sxs-lookup"><span data-stu-id="e50e6-124">HTML</span></span> | <span data-ttu-id="e50e6-125">提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。</span><span class="sxs-lookup"><span data-stu-id="e50e6-125">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="e50e6-126">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="e50e6-126">**./manifest.xml**</span></span> | <span data-ttu-id="e50e6-127">XML</span><span class="sxs-lookup"><span data-stu-id="e50e6-127">XML</span></span> | <span data-ttu-id="e50e6-128">指定自定义函数使用的多个文件的位置，例如自定义函数 JavaScript、JSON 和 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="e50e6-128">Specifies the location of multiple files that your custom function use, such as the custom functions JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="e50e6-129">它还列出了任务窗格文件、命令文件的位置，并指定自定义函数应使用的运行时。</span><span class="sxs-lookup"><span data-stu-id="e50e6-129">It also lists the locations of task pane files, command files, and specifies which runtime your custom functions should use.</span></span> |

### <a name="script-file"></a><span data-ttu-id="e50e6-130">脚本文件</span><span class="sxs-lookup"><span data-stu-id="e50e6-130">Script file</span></span>

<span data-ttu-id="e50e6-131">脚本文件 (**./src/functions/functions.js** or **./src/functions/functions.ts**) 包含定义自定义函数的代码以及定义函数的注释。</span><span class="sxs-lookup"><span data-stu-id="e50e6-131">The script file (**./src/functions/functions.js** or **./src/functions/functions.ts**) contains the code that defines custom functions and comments which define the function.</span></span>

<span data-ttu-id="e50e6-132">以下代码定义 `add` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="e50e6-132">The following code defines the custom function `add`.</span></span> <span data-ttu-id="e50e6-133">代码注释用于生成描述 Excel 自定义函数的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="e50e6-133">The code comments are used to generate a JSON metadata file that describes the custom function to Excel.</span></span> <span data-ttu-id="e50e6-134">首先声明所需的 `@customfunction` 注释，指示这是一个自定义函数。</span><span class="sxs-lookup"><span data-stu-id="e50e6-134">The required `@customfunction` comment is declared first, to indicate that this is a custom function.</span></span> <span data-ttu-id="e50e6-135">接下来，声明两个参数 `first` 和 `second`，然后是它们的 `description` 属性。</span><span class="sxs-lookup"><span data-stu-id="e50e6-135">Next, two parameters are declared, `first` and `second`, followed by their `description` properties.</span></span> <span data-ttu-id="e50e6-136">最后提供了 `returns` 描述。</span><span class="sxs-lookup"><span data-stu-id="e50e6-136">Finally, a `returns` description is given.</span></span> <span data-ttu-id="e50e6-137">要详细了解自定义函数需要哪些注释，请参阅[为自定义函数创建 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="e50e6-137">For more information about what comments are required for your custom function, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

### <a name="manifest-file"></a><span data-ttu-id="e50e6-138">清单文件</span><span class="sxs-lookup"><span data-stu-id="e50e6-138">Manifest file</span></span>

<span data-ttu-id="e50e6-139">用于定义自定义函数的加载项的 XML 清单文件（Yo Office 生成器创建的项目中的 **./manifest.xml**）会执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="e50e6-139">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) does several things:</span></span>

- <span data-ttu-id="e50e6-140">定义自定义函数的命名空间。</span><span class="sxs-lookup"><span data-stu-id="e50e6-140">Defines the namespace for your custom functions.</span></span> <span data-ttu-id="e50e6-141">命名空间追加在你的自定义函数之前，可帮助客户将你的函数标识为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="e50e6-141">A namespace prepends itself to your custom functions to help customers identify your functions as part of your add-in.</span></span>
- <span data-ttu-id="e50e6-142">使用自定义函数清单特有的 `<ExtensionPoint>` 和 `<Resources>` 元素。</span><span class="sxs-lookup"><span data-stu-id="e50e6-142">Uses `<ExtensionPoint>` and `<Resources>` elements that are unique to a custom functions manifest.</span></span> <span data-ttu-id="e50e6-143">这些元素包含有关 JavaScript、JSON 和 HTML 文件的位置的信息。</span><span class="sxs-lookup"><span data-stu-id="e50e6-143">These elements contain the information about the locations of the JavaScript, JSON, and HTML files.</span></span>
- <span data-ttu-id="e50e6-144">指定要用于自定义函数的运行时。</span><span class="sxs-lookup"><span data-stu-id="e50e6-144">Specifies which runtime to use for your custom function.</span></span> <span data-ttu-id="e50e6-145">除非你对另一运行时有特殊需求，否则建议始终使用共享运行时，因为共享运行时允许在函数和任务窗格之间共享数据。</span><span class="sxs-lookup"><span data-stu-id="e50e6-145">We recommend always using a shared runtime unless you have a specific need for another runtime, because a shared runtime allows for the sharing of data between functions and the task pane.</span></span> <span data-ttu-id="e50e6-146">请注意，使用共享运行时意味着加载项将使用 Internet Explorer 11，而不是 Microsoft Edge。</span><span class="sxs-lookup"><span data-stu-id="e50e6-146">Note that using a shared runtime means your add-in will use Internet Explorer 11, not Microsoft Edge.</span></span>

<span data-ttu-id="e50e6-147">如果你使用 Yo Office 生成器来创建文件，则建议将你的清单调整为使用共享运行时，因为这不是这些文件的默认设置。</span><span class="sxs-lookup"><span data-stu-id="e50e6-147">If you are using the Yo Office generator to create files, we recommend adjusting your manifest to use a shared runtime, as this is not the default for these files.</span></span> <span data-ttu-id="e50e6-148">若要更改清单，请按照[将 Excel 加载项配置为使用共享 JavaScript 运行时](configure-your-add-in-to-use-a-shared-runtime.md)中的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="e50e6-148">To change your manifest, follow the instructions in [Configure your Excel add-in to use a shared JavaScript runtime](configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="e50e6-149">若要从示例加载项查看完整的工作清单，请参阅[此 Github 存储库](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml)。</span><span class="sxs-lookup"><span data-stu-id="e50e6-149">To see a full working manifest from a sample add-in, see [this Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a><span data-ttu-id="e50e6-150">共同创作</span><span class="sxs-lookup"><span data-stu-id="e50e6-150">Coauthoring</span></span>

<span data-ttu-id="e50e6-151">利用连接到 Microsoft 365 订阅的 Excel web 版和 Windows 版 Excel，你可以在 Excel 中共同创作。</span><span class="sxs-lookup"><span data-stu-id="e50e6-151">Excel on the web and Windows connected to a Microsoft 365 subscription allow you to coauthor in Excel.</span></span> <span data-ttu-id="e50e6-152">如果你的工作簿使用自定义函数，系统会提示你的共同创作同事加载自定义函数的加载项。</span><span class="sxs-lookup"><span data-stu-id="e50e6-152">If your workbook uses a custom function, your coauthoring colleague is prompted to load the custom function's add-in.</span></span> <span data-ttu-id="e50e6-153">当你们均加载此加载项后，自定义函数将通过共同创作共享结果。</span><span class="sxs-lookup"><span data-stu-id="e50e6-153">Once you both have loaded the add-in, the custom function shares results through coauthoring.</span></span>

<span data-ttu-id="e50e6-154">若要详细了解共同创作，请参阅[关于 Excel 中的共同创作](/office/vba/excel/concepts/about-coauthoring-in-excel)。</span><span class="sxs-lookup"><span data-stu-id="e50e6-154">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="known-issues"></a><span data-ttu-id="e50e6-155">已知问题</span><span class="sxs-lookup"><span data-stu-id="e50e6-155">Known issues</span></span>

<span data-ttu-id="e50e6-156">在 [Excel 自定义功能 GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/issues)上查看已知问题。</span><span class="sxs-lookup"><span data-stu-id="e50e6-156">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="next-steps"></a><span data-ttu-id="e50e6-157">后续步骤</span><span class="sxs-lookup"><span data-stu-id="e50e6-157">Next steps</span></span>

<span data-ttu-id="e50e6-158">想要试用自定义函数？</span><span class="sxs-lookup"><span data-stu-id="e50e6-158">Want to try out custom functions?</span></span> <span data-ttu-id="e50e6-159">检查简单的[自定义函数入门](../quickstarts/excel-custom-functions-quickstart.md)或更深入的[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)（如果还没有）。</span><span class="sxs-lookup"><span data-stu-id="e50e6-159">Check out the simple [custom functions quick start](../quickstarts/excel-custom-functions-quickstart.md) or the more in-depth [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md) if you haven't already.</span></span>

<span data-ttu-id="e50e6-160">另一个尝试自定义函数的简单方法就是使用[脚本实验室](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)，这是一个允许您在 Excel 中试验自定义函数的加载项。</span><span class="sxs-lookup"><span data-stu-id="e50e6-160">Another easy way to try out custom functions is to use [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), an add-in that allows you to experiment with custom functions right in Excel.</span></span> <span data-ttu-id="e50e6-161">可以尝试创建自己的自定义函数或使用提供的示例。</span><span class="sxs-lookup"><span data-stu-id="e50e6-161">You can try out creating your own custom function or play with the provided samples.</span></span>

## <a name="see-also"></a><span data-ttu-id="e50e6-162">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e50e6-162">See also</span></span> 
* [<span data-ttu-id="e50e6-163">自定义函数要求</span><span class="sxs-lookup"><span data-stu-id="e50e6-163">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="e50e6-164">命名准则</span><span class="sxs-lookup"><span data-stu-id="e50e6-164">Naming guidelines</span></span>](custom-functions-naming.md)
* [<span data-ttu-id="e50e6-165">让自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="e50e6-165">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
