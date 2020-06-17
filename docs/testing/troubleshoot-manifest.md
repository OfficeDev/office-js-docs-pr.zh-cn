---
title: 验证 Office 加载项的清单
description: 了解如何使用 XML 架构和其他工具验证 Office 外接程序的清单。
ms.date: 04/16/2020
localization_priority: Normal
ms.openlocfilehash: fee4fd048092734eb479f1993c69fcf99c153c79
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611097"
---
# <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="f5756-103">验证 Office 加载项的清单</span><span class="sxs-lookup"><span data-stu-id="f5756-103">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="f5756-104">你可能需要验证加载项的清单文件，以确保其正确且完整。</span><span class="sxs-lookup"><span data-stu-id="f5756-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="f5756-105">当你尝试旁加载加载项时，验证还可以识别导致错误“你的加载项清单无效”的问题。</span><span class="sxs-lookup"><span data-stu-id="f5756-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="f5756-106">本文介绍了验证清单文件的多种方法。</span><span class="sxs-lookup"><span data-stu-id="f5756-106">This article describes multiple ways to validate the manifest file.</span></span>

> [!NOTE]
> <span data-ttu-id="f5756-107">有关使用运行时日志记录功能来解决加载项清单问题的详细信息，请参阅[使用运行时日志记录功能调试加载项](runtime-logging.md)。</span><span class="sxs-lookup"><span data-stu-id="f5756-107">For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="f5756-108">使用适用于 Office 加载项的 Yeoman 生成器来验证清单</span><span class="sxs-lookup"><span data-stu-id="f5756-108">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="f5756-109">如果你使用了[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建加载项，则也可以使用它来验证项目的清单文件。</span><span class="sxs-lookup"><span data-stu-id="f5756-109">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="f5756-110">在项目的根目录中运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="f5756-110">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![动画 gif 显示 Yo Office 验证程序（在命令行处运行并生成显示“验证已通过”的结果）。](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="f5756-112">若要访问此功能，必须使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)版本 1.1.17 或更高版本创建加载项项目。</span><span class="sxs-lookup"><span data-stu-id="f5756-112">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="f5756-113">使用 office-addin-manifest 验证清单</span><span class="sxs-lookup"><span data-stu-id="f5756-113">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="f5756-114">如果你未使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建加载项，则可以使用 [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest)。</span><span class="sxs-lookup"><span data-stu-id="f5756-114">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="f5756-115">安装 [Node.js](https://nodejs.org/download/)。</span><span class="sxs-lookup"><span data-stu-id="f5756-115">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="f5756-116">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="f5756-116">Run the following command in the root directory of your project.</span></span> 

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > <span data-ttu-id="f5756-117">如果此命令不可用或不起作用，请运行以下命令来强制使用 office 外接程序清单工具的最新版本（替换 `MANIFEST_FILE` 为清单文件的名称）：</span><span class="sxs-lookup"><span data-stu-id="f5756-117">If this command is not available or not working, run the following command instead to force the use of the latest version of the office-addin-manifest tool (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span>
    >
    > ```command&nbsp;line
    > npx --ignore-existing office-addin-manifest validate MANIFEST_FILE
    > ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="f5756-118">根据 XML 架构验证清单</span><span class="sxs-lookup"><span data-stu-id="f5756-118">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="f5756-119">可以根据 [XML 架构定义 (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) 文件来验证清单文件。</span><span class="sxs-lookup"><span data-stu-id="f5756-119">You can validate the manifest file against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) files.</span></span> <span data-ttu-id="f5756-120">这将有助于确保清单文件采用正确架构（包括所使用的元素的所有命名空间）。</span><span class="sxs-lookup"><span data-stu-id="f5756-120">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="f5756-121">如果从其他示例清单中复制了元素，请仔细检查是否还**添加了相应命名空间**。</span><span class="sxs-lookup"><span data-stu-id="f5756-121">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="f5756-122">若要执行此验证，可以使用 XML 架构验证工具。</span><span class="sxs-lookup"><span data-stu-id="f5756-122">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="f5756-123">使用命令行 XML 架构验证工具验证清单的具体步骤</span><span class="sxs-lookup"><span data-stu-id="f5756-123">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="f5756-124">安装 [tar](https://www.gnu.org/software/tar/) 和 [libxml](http://xmlsoft.org/FAQ.html)（如果尚未安装的话）。</span><span class="sxs-lookup"><span data-stu-id="f5756-124">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="f5756-p104">运行下面的命令。将 `XSD_FILE` 替换为清单 XSD 文件路径，并将 `XML_FILE` 替换为清单 XML 文件路径。</span><span class="sxs-lookup"><span data-stu-id="f5756-p104">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a><span data-ttu-id="f5756-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f5756-127">See also</span></span>

- [<span data-ttu-id="f5756-128">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="f5756-128">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="f5756-129">清除 Office 缓存</span><span class="sxs-lookup"><span data-stu-id="f5756-129">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="f5756-130">使用运行时日志记录功能调试加载项</span><span class="sxs-lookup"><span data-stu-id="f5756-130">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="f5756-131">旁加载 Office 外接程序进行测试</span><span class="sxs-lookup"><span data-stu-id="f5756-131">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="f5756-132">调试 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="f5756-132">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
