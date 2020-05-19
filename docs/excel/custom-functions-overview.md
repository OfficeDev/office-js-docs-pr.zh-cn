---
ms.date: 05/17/2020
description: 为 Office 加载项创建 Excel 自定义函数
title: 在 Excel 中创建自定义函数
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: dabb196bc4b55bd4852f9c857767dcabd3063045
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44276006"
---
# <a name="create-custom-functions-in-excel"></a>在 Excel 中创建自定义函数

开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。 Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

以下动态图像显示调用你使用 JavaScript 或 Typescript 创建的函数的工作簿。 在此示例中，自定义函数 `=MYFUNCTION.SPHEREVOLUME` 计算球的体积。

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

以下代码定义 `=MYFUNCTION.SPHEREVOLUME` 自定义函数。

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
> 本文后面的[已知问题](#known-issues)部分指定自定义函数的当前限制。

## <a name="how-a-custom-function-is-defined-in-code"></a>如何在代码中定义自定义函数

如果使用[Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，它将创建用于控制函数和任务窗格的文件。 我们将专注于对自定义函数至关重要的文件：

| 文件 | 文件格式 | 说明 |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>或<br/>**./src/functions/functions.ts** | JavaScript<br/>或<br/>TypeScript | 包含定义自定义函数的代码。 |
| **./src/functions/functions.html** | HTML | 提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。 |
| **./manifest.xml** | XML | 指定自定义函数使用的多个文件的位置，例如自定义函数 JavaScript、JSON 和 HTML 文件。 此外，它还列出了任务窗格文件、命令文件的位置，并指定了自定义函数应使用的运行时。 |

### <a name="script-file"></a>脚本文件

脚本文件 (**./src/functions/functions.js** or **./src/functions/functions.ts**) 包含定义自定义函数的代码以及定义函数的注释。

以下代码定义 `add` 自定义函数。 代码注释用于生成描述 Excel 自定义函数的 JSON 元数据。 首先声明所需的 `@customfunction` 注释，指示这是一个自定义函数。 接下来，先声明两个参数， `first` 然后再 `second` 键入它们的 `description` 属性。 最后提供了 `returns` 描述。 要详细了解自定义函数需要哪些注释，请参阅[为自定义函数创建 JSON 元数据](custom-functions-json-autogeneration.md)。

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

### <a name="manifest-file"></a>清单文件

用于定义自定义函数（在 Yo Office 生成器创建的项目中的 **/manifest.xml** ）的外接程序的 XML 清单文件执行以下操作：

- 定义自定义函数的命名空间。 命名空间将自己添加到自定义函数中，以帮助客户将您的函数标识为外接程序的一部分。
- 使用 `<ExtensionPoint>` 和 `<Resources>` 元素对于自定义函数清单而言是唯一的。 这些元素包含有关 JavaScript、JSON 和 HTML 文件的位置的信息。
- 指定要用于自定义函数的运行时。 我们建议始终使用共享运行时，除非您有其他运行时的特定需求，因为共享运行时允许在函数和任务窗格之间共享数据。

如果使用 Yo Office 生成器创建文件，建议将清单调整为使用共享运行时，因为这不是这些文件的默认值。 若要更改清单，请按照[配置 Excel 外接程序中的说明使用共享的 JavaScript 运行时](./configure-your-add-in-to-use-a-shared-runtime.md)。

若要查看示例加载项中的完整工作清单，请参阅[此 Github 存储库](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml)。

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a>共同创作

Web 上的 Excel 和连接到 Office 365 订阅的 Windows 允许您在 Excel 中 coauthor。 如果您的工作簿使用自定义函数，则将提示您的合著同事加载自定义函数的外接程序。 一旦您加载了加载项，自定义函数将通过共同创作来共享结果。

若要详细了解共同创作，请参阅[关于 Excel 中的共同创作](/office/vba/excel/concepts/about-coauthoring-in-excel)。

## <a name="known-issues"></a>已知问题

在 [Excel 自定义功能 GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/issues)上查看已知问题。

## <a name="next-steps"></a>后续步骤

想要试用自定义函数？ 检查简单的[自定义函数入门](../quickstarts/excel-custom-functions-quickstart.md)或更深入的[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)（如果还没有）。

另一个尝试自定义函数的简单方法就是使用[脚本实验室](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)，这是一个允许您在 Excel 中试验自定义函数的加载项。 可以尝试创建自己的自定义函数或使用提供的示例。

## <a name="see-also"></a>另请参阅 
* [自定义函数要求](custom-functions-requirement-sets.md)
* [命名准则](custom-functions-naming.md)
* [让自定义函数与 XLL 用户定义的函数兼容](make-custom-functions-compatible-with-xll-udf.md)
