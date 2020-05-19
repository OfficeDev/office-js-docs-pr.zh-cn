---
ms.date: 04/29/2020
description: 本地化您的 Excel 自定义函数。
title: 本地化自定义函数
localization_priority: Normal
ms.openlocfilehash: 001045f82634d7e96c4d4515ccd87b5cfaf2cd1c
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275964"
---
# <a name="localize-custom-functions"></a><span data-ttu-id="ac8a3-103">本地化自定义函数</span><span class="sxs-lookup"><span data-stu-id="ac8a3-103">Localize custom functions</span></span>

<span data-ttu-id="ac8a3-104">您可以本地化您的外接程序和自定义函数名称。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-104">You can localize both your add-in and your custom function names.</span></span> <span data-ttu-id="ac8a3-105">若要执行此操作，请在函数的 JSON 文件中提供本地化的函数名称，并在 XML 清单文件中提供区域设置信息。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-105">To do so, provide localized function names in the functions' JSON file and locale information in the XML manifest file.</span></span>

>[!IMPORTANT]
> <span data-ttu-id="ac8a3-106">自动生成的元数据不能用于本地化，因此您需要手动更新 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-106">Auto-generated metadata doesn't work for localization so you need to update the JSON file manually.</span></span> <span data-ttu-id="ac8a3-107">若要了解如何执行此操作，请参阅[Excel 中的自定义函数的元数据](custom-functions-json.md)</span><span class="sxs-lookup"><span data-stu-id="ac8a3-107">To learn how to do this, see [Metadata for custom functions in Excel](custom-functions-json.md)</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a><span data-ttu-id="ac8a3-108">本地化函数名称</span><span class="sxs-lookup"><span data-stu-id="ac8a3-108">Localize function names</span></span>

<span data-ttu-id="ac8a3-109">若要本地化自定义函数，请为每种语言创建一个新的 JSON 元数据文件。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-109">To localize your custom functions, create a new JSON metadata file for each language.</span></span> <span data-ttu-id="ac8a3-110">在每个语言 JSON 文件中 `name` ， `description` 在目标语言中创建和属性。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-110">In each language JSON file, create `name` and `description` properties in the target language.</span></span> <span data-ttu-id="ac8a3-111">英语的默认文件命名为**函数 json**。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-111">The default file for English is named **functions.json**.</span></span> <span data-ttu-id="ac8a3-112">对每个其他 JSON 文件使用文件名中的区域设置（如**函数-** 为帮助识别它们）。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-112">Use the locale in the filename for each additional JSON file, such as **functions-de.json** to help identify them.</span></span>

<span data-ttu-id="ac8a3-113">`name`并 `description` 将显示在 Excel 中并进行本地化。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-113">The `name` and `description` appear in Excel and are localized.</span></span> <span data-ttu-id="ac8a3-114">但是，不会对 `id` 每个函数的进行本地化。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-114">However, the `id` of each function isn't localized.</span></span> <span data-ttu-id="ac8a3-115">`id`属性是 Excel 将函数标识为唯一的，不应在设置后更改。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-115">The `id` property is how Excel identifies your function as unique and shouldn't be changed once it is set.</span></span>

<span data-ttu-id="ac8a3-116">下面的 JSON 演示如何定义 `id` 属性 "乘法" 的函数。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-116">The following JSON shows how to define a function with the `id` property "MULTIPLY."</span></span> <span data-ttu-id="ac8a3-117">`name`函数的和 `description` 属性本地化为德语。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-117">The `name` and `description` property of the function is localized for German.</span></span> <span data-ttu-id="ac8a3-118">每个 `name` 参数 `description` 也本地化为德语。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-118">Each parameter `name` and `description` is also localized for German.</span></span>

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

<span data-ttu-id="ac8a3-119">将以前的 JSON 与以下 JSON 进行比较，以获取英语。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-119">Compare the previous JSON with the following JSON for English.</span></span>

```JSON
{
    "id": "MULTIPLY",
    "name": "Multiply",
    "description": "Multiplies two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "one",
            "description": "first number",
            "dimensionality": "scalar"
        },
        {
            "name": "two",
            "description": "second number",
            "dimensionality": "scalar"
        },
    ],
}
```

## <a name="localize-your-add-in"></a><span data-ttu-id="ac8a3-120">本地化你的外接程序</span><span class="sxs-lookup"><span data-stu-id="ac8a3-120">Localize your add-in</span></span>

<span data-ttu-id="ac8a3-121">为每种语言创建一个 JSON 文件后，使用指定每个 JSON 元数据文件的 URL 的每个区域设置的替代值更新 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-121">After creating a JSON file for each language, update your XML manifest file with an override value for each locale that specifies the URL of each JSON metadata file.</span></span> <span data-ttu-id="ac8a3-122">下面的清单 XML 显示了一个默认 `en-us` 区域设置，其中包含 `de-de` （德国）的覆盖 JSON 文件 URL。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-122">The following manifest XML shows a default `en-us` locale with an override JSON file URL for `de-de` (Germany).</span></span> <span data-ttu-id="ac8a3-123">**函数-. json**文件包含本地化的德语函数名称和 id。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-123">The **functions-de.json** file contains the localized German function names and ids.</span></span>

```XML
<DefaultLocale>en-us</DefaultLocale>
...
<Resources>
     <bt:Urls>
        <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
          <bt:Override Locale="de-de" Value="https://localhost:3000/dist/functions-de.json" />
        </bt:url>
        
     </bt:Urls>
</Resources>
```

<span data-ttu-id="ac8a3-124">有关本地化外接程序的过程的详细信息，请参阅[Office 外接程序的本地化](../develop/localization.md#control-localization-from-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-124">For more information on the process of localizing an add-in, see [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).</span></span>

## <a name="next-steps"></a><span data-ttu-id="ac8a3-125">后续步骤</span><span class="sxs-lookup"><span data-stu-id="ac8a3-125">Next steps</span></span>
<span data-ttu-id="ac8a3-126">了解[自定义函数的命名约定](custom-functions-naming.md)或发现[错误处理最佳实践](custom-functions-errors.md)。</span><span class="sxs-lookup"><span data-stu-id="ac8a3-126">Learn about [naming conventions for custom functions](custom-functions-naming.md) or discover [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ac8a3-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ac8a3-127">See also</span></span>

* [<span data-ttu-id="ac8a3-128">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="ac8a3-128">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ac8a3-129">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="ac8a3-129">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="ac8a3-130">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="ac8a3-130">Create custom functions in Excel</span></span>](custom-functions-overview.md)
