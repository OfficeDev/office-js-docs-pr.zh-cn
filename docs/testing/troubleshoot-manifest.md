---
title: 验证 Office 加载项的清单
description: 了解如何使用 XML 架构和其他工具Office外接程序的清单。
ms.date: 10/29/2020
ms.localizationpriority: medium
ms.openlocfilehash: 30e7b93430b8ddffc5ebc2cc8f2ae2bab5c0850f
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681526"
---
# <a name="validate-an-office-add-ins-manifest"></a>验证 Office 加载项的清单

你可能需要验证加载项的清单文件，以确保其正确且完整。 当你尝试旁加载加载项时，验证还可以识别导致错误“你的加载项清单无效”的问题。 本文介绍了验证清单文件的多种方法。

> [!NOTE]
> 有关使用运行时日志记录功能来解决加载项清单问题的详细信息，请参阅[使用运行时日志记录功能调试加载项](runtime-logging.md)。

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>使用适用于 Office 加载项的 Yeoman 生成器来验证清单

如果你使用了[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建加载项，则也可以使用它来验证项目的清单文件。 在项目的根目录中运行以下命令。

```command&nbsp;line
npm run validate
```

![动态 GIF，显示 Yo Office验证程序在命令行中运行并生成显示"验证通过"的结果。](../images/yo-office-validator.gif)

> [!NOTE]
> 若要访问此功能，必须使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)版本 1.1.17 或更高版本创建加载项项目。

## <a name="validate-your-manifest-with-office-addin-manifest"></a>使用 office-addin-manifest 验证清单

如果你未使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建加载项，则可以使用 [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest)。

1. 安装 [Node.js](https://nodejs.org/download/)。

1. 打开命令提示符，然后使用以下命令安装验证程序。

    ```command&nbsp;line
    npm install -g office-addin-manifest
    ```

1. 在项目的 *根目录中运行以下命令*。

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > 如果此命令不可用或无法工作，请改为运行以下命令以强制使用最新版本的 office-addin-manifest 工具 (以清单文件名称替换 `MANIFEST_FILE`) 。
    >
    > ```command&nbsp;line
    > npx office-addin-manifest validate MANIFEST_FILE
    > ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>根据 XML 架构验证清单

可以根据 [XML 架构定义 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) 文件来验证清单文件。 这将有助于确保清单文件采用正确架构（包括所使用的元素的所有命名空间）。 如果从其他示例清单中复制了元素，请仔细检查是否还 **添加了相应命名空间**。 若要执行此验证，可以使用 XML 架构验证工具。

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>使用命令行 XML 架构验证工具验证清单的具体步骤

1. 安装 [tar](https://www.gnu.org/software/tar/) 和 [libxml](http://xmlsoft.org/FAQ.html)（如果尚未安装的话）。

1. 运行下面的命令。将 `XSD_FILE` 替换为清单 XSD 文件路径，并将 `XML_FILE` 替换为清单 XML 文件路径。

    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](../develop/add-in-manifests.md)
- [清除 Office 缓存](clear-cache.md)
- [使用运行时日志记录功能调试加载项](runtime-logging.md)
- [旁加载 Office 外接程序进行测试](sideload-office-add-ins-for-testing.md)
- [使用适用于 Internet Explorer 的开发人员工具调试加载项](debug-add-ins-using-f12-tools-ie.md)
- [使用旧版 Edge 开发人员工具调试加载项](debug-add-ins-using-devtools-edge-legacy.md)
