---
description: 为 Office 加载项创建 Excel 自定义函数。
title: 在 Excel 中创建自定义函数
ms.date: 08/04/2021
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 12740615215913b0201426f929dbcb803c866648
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422759"
---
# <a name="create-custom-functions-in-excel"></a>在 Excel 中创建自定义函数

开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。 Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

以下动态图像显示调用你使用 JavaScript 或 TypeScript 创建的函数的工作簿。 在此示例中，自定义函数 `=MYFUNCTION.SPHEREVOLUME` 计算球的体积。

![显示最终用户插入 MYFUNCTION 的动画图像。将 SPHEREVOLUME 自定义函数放入 Excel 工作表的单元格中。](../images/SphereVolumeNew.gif)

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

> [!TIP]
> 如果自定义函数外接程序将使用任务窗格或功能区按钮，除了运行自定义函数代码外，还需要设置 [共享运行时](../testing/runtimes.md#shared-runtime)。 若要了解详细信息，请参阅 [配置 Office 加载项以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="how-a-custom-function-is-defined-in-code"></a>如何在代码中定义自定义函数

如果使用 [Yeoman 的 office 外接程序生成器](../develop/yeoman-generator-overview.md) 创建 Excel 自定义函数外接程序项目，它将创建控制函数和任务窗格的文件。我们将重点介绍对自定义函数重要的文件。

| 文件 | 文件格式 | 说明 |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>或<br/>**./src/functions/functions.ts** | JavaScript<br/>或<br/>TypeScript | 包含定义自定义函数的代码。 |
| **./src/functions/functions.html** | HTML | 提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。 |
| **./manifest.xml** | XML | 指定自定义函数使用的多个文件的位置，例如自定义函数 JavaScript、JSON 和 HTML 文件。 它还列出了任务窗格文件、命令文件的位置，并指定自定义函数应使用的运行时。 |

### <a name="script-file"></a>脚本文件

脚本文件 (**./src/functions/functions.js** or **./src/functions/functions.ts**) 包含定义自定义函数的代码以及定义函数的注释。

以下代码定义 `add` 自定义函数。 代码注释用于生成描述 Excel 自定义函数的 JSON 元数据。 首先声明所需的 `@customfunction` 注释，指示这是一个自定义函数。 接下来，声明两个参数 `first` 和 `second`，然后是它们的 `description` 属性。 最后提供了 `returns` 描述。 要详细了解自定义函数需要哪些注释，请参阅[为自定义函数创建 JSON 元数据](custom-functions-json-autogeneration.md)。

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

用于定义自定义函数的加载项的 XML 清单文件（[适用于 Office 的 Yeoman 生成器](../develop/yeoman-generator-overview.md) 创建的项目中的 **./manifest.xml**）会执行以下操作。

- 定义自定义函数的命名空间。命名空间在自定义函数前加上自己的名字，可帮助客户识别加载项的的函数。
- 使用自定义函数清单特有的 **\<ExtensionPoint\>** 和 **\<Resources\>** 元素。 这些元素包含有关 JavaScript、JSON 和 HTML 文件的位置的信息。
- 指定用于自定义函数的运行时。除非你对另一运行时有特殊需求，否则建议始终使用共享运行时，因为共享运行时允许在函数和任务窗格之间共享数据。

如果使用 [Office 外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md) 创建文件，建议调整清单以使用共享运行时，因为这不是这些文件的默认值。 若要更改清单，请按照配置 [Excel 外接程序中的说明使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

要从示例加载项中查看完整的工作清单，请参阅 [我们的 Office 加载项示例 Github 存储库之一](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Samples/excel-shared-runtime-global-state/manifest.xml) 中的清单。

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a>共同创作

利用连接到 Microsoft 365 订阅的 Excel 网页版和 Windows 版 Excel，最终用户可以在 Excel 中共同创作。 如果最终用户的工作簿使用自定义函数，系统将提示该最终用户的共同创作同事加载相应的自定义函数加载项。 当两个用户均加载此加载项后，自定义函数将通过共同创作共享结果。

若要详细了解共同创作，请参阅[关于 Excel 中的共同创作](/office/vba/excel/concepts/about-coauthoring-in-excel)。

## <a name="next-steps"></a>后续步骤

想要试用自定义函数？请查看简单的 [自定义函数快速入门](../quickstarts/excel-custom-functions-quickstart.md) 或更深入的 [自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)（如果尚未这样做）。

另一个尝试自定义函数的简单方法就是使用[脚本实验室](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)，这是一个允许您在 Excel 中试验自定义函数的加载项。 可以尝试创建自己的自定义函数或使用提供的示例。

## <a name="see-also"></a>另请参阅

- [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
- [自定义函数要求集](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)
- [自定义函数命名准则](custom-functions-naming.md)
- [让自定义函数与 XLL 用户定义的函数兼容](make-custom-functions-compatible-with-xll-udf.md)
- [将 Office 外接程序配置为使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Office 加载项中的运行时](../testing/runtimes.md)
