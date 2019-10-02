---
ms.date: 09/26/2019
description: 在 Excel 中使用 JavaScript 创建自定义函数。
title: 在 Excel 中创建自定义函数
ms.topic: overview
scenarios: getting-started
localization_priority: Priority
ms.openlocfilehash: aeff31dfce62aa9983f9a4ff2d766e127314d967
ms.sourcegitcommit: 528577145b2cf0a42bc64c56145d661c4d019fb8
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/02/2019
ms.locfileid: "37353865"
---
# <a name="create-custom-functions-in-excel"></a>在 Excel 中创建自定义函数 

开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。 Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。 本文介绍了如何在 Excel 中创建自定义函数。

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

如果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，会发现它可创建全面控制函数、任务窗格和加载项的文件。 我们将专注于对自定义函数至关重要的文件：

| 文件 | 文件格式 | 说明 |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>或<br/>**./src/functions/functions.ts** | JavaScript<br/>或<br/>TypeScript | 包含定义自定义函数的代码。 |
| **./src/functions/functions.html** | HTML | 提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。 |
| **./manifest.xml** | XML | 指定加载项中所有自定义函数的命名空间以及此表中前面列出的 JavaScript 和 HTML 文件的位置。 它还列出了加载项可能使用的其他文件的位置，如任务窗格文件和命令文件。 |

### <a name="script-file"></a>脚本文件

脚本文件 (**./src/functions/functions.js** or **./src/functions/functions.ts**) 包含定义自定义函数的代码以及定义函数的注释。

以下代码定义 `add` 自定义函数。 代码注释用于生成描述 Excel 自定义函数的 JSON 元数据。 首先声明所需的 `@customfunction` 注释，指示这是一个自定义函数。 此外，你将注意到声明了两个参数，即 `first` 和 `second`，后跟其 `description` 属性。 最后提供了 `returns` 描述。 要详细了解自定义函数需要哪些注释，请参阅[为自定义函数创建 JSON 元数据](custom-functions-json-autogeneration.md)。

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

请注意，控制自定义函数运行时加载的 **functions.html** 文件必须链接至自定义函数的当前 CDN。 准备有当前版本的 Yo Office 生成器的项目引用正确的 CDN。 如果更新 2019 年 3 月或更早的自定义函数项目，则需要将以下代码复制到 **functions.html** 页面。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/custom-functions-runtime.js" type="text/javascript"></script>
```

### <a name="manifest-file"></a>清单文件

定义自定义函数的加载项的 XML 清单文件（Yo Office 生成器创建的项目中的 **./manifest.xml**）指定加载项中所有自定义函数的命名空间以及 JavaScript、JSON 和 HTML 文件的位置。

下面的基本 XML 标记显示了 `<ExtensionPoint>` 和 `<Resources>` 元素的一个示例，必须在加载项清单中包含这些元素才能启用自定义函数。 如果使用 Yo Office 生成器，生成的自定义函数文件将包含更复杂的清单文件，可以在[此 Github 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml)中对其进行比较。

> [!NOTE] 
> 在自定义函数 JavaScript、JSON 和 HTML 文件的清单文件中指定的 URL 必须可公开访问，并具有相同的子域。

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

## <a name="coauthoring"></a>共同创作

借助已连接到 Office 365 订阅的 Excel 网页版和 Windows 版 Excel，可以共同创作文档；此功能可与自定义函数结合使用。 如果你的工作簿使用自定义函数，系统会提示你的同事加载自定义函数的加载项。 当你们均加载此加载项后，自定义函数会通过共同创作共享结果。

若要详细了解共同创作，请参阅[关于 Excel 中的共同创作](/office/vba/excel/concepts/about-coauthoring-in-excel)。

## <a name="known-issues"></a>已知问题

在 [Excel 自定义功能 GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/issues)上查看已知问题。

## <a name="next-steps"></a>后续步骤

想要试用自定义函数？ 检查简单的[自定义函数入门](../quickstarts/excel-custom-functions-quickstart.md)或更深入的[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)（如果还没有）。

另一个尝试自定义函数的简单方法就是使用[脚本实验室](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab)，这是一个允许您在 Excel 中试验自定义函数的加载项。 可以尝试创建自己的自定义函数或使用提供的示例。

准备详细了解自定义函数的功能？ 了解[自定义函数架构](custom-functions-architecture.md)的概述。

## <a name="see-also"></a>另请参阅 
* [自定义函数要求](custom-functions-requirement-sets.md)
* [命名准则](custom-functions-naming.md)
* [让自定义函数与 XLL 用户定义的函数兼容](make-custom-functions-compatible-with-xll-udf.md)
