---
ms.date: 11/06/2020
description: 本地化自定义Excel函数。
title: 本地化自定义函数
ms.localizationpriority: medium
ms.openlocfilehash: 7219c838cfd5a6c827b74b5d04442280be7ebac7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744504"
---
# <a name="localize-custom-functions"></a>本地化自定义函数

您可以本地化外接程序和自定义函数名称。 为此，在函数的 JSON 文件中提供本地化函数名称，在 XML 清单文件中提供区域设置信息。

>[!IMPORTANT]
> 自动生成的元数据不能用于本地化，因此你需要手动更新 JSON 文件。 若要了解如何操作，请参阅手动为自定义函数 [创建 JSON 元数据](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>本地化函数名称

若要本地化自定义函数，请为每个语言创建新的 JSON 元数据文件。 在每个语言 JSON 文件中，使用`name``description`目标语言创建 和 属性。 英语的默认文件名为 **functions.json**。 使用文件名中每个其他 JSON 文件（如 **functions-de.json** ）中的区域设置来帮助识别它们。

和 `name` 显示在 `description` Excel中，并本地化。 但是， `id` 不会本地化每个函数的 。 属性`id`是Excel函数的唯一性，且设置后不应更改它。

以下 JSON 显示如何使用属性 `id` "MULTIPLY"定义函数。 `description`函数`name`的 和 属性针对德语进行本地化。 每个参数 `name` 和 `description` 还针对德语进行本地化。

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

比较以前的 JSON 和以下 JSON 表示英语。

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

## <a name="localize-your-add-in"></a>本地化外接程序

为每种语言创建 JSON 文件后，使用每个区域设置（用于指定每个 JSON 元数据文件的 URL）的替代值更新 XML 清单文件。 以下清单 XML 显示了一个默认 `en-us` 区域设置，其替代了德国 (JSON `de-de`) 。 **functions-de.json** 文件包含本地化的德语函数名称和 ID。

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

有关本地化外接程序的过程详细信息，请参阅本地化 Office [外接程序](../develop/localization.md#control-localization-from-the-manifest)。

## <a name="next-steps"></a>后续步骤
了解 [自定义函数的命名约定或](custom-functions-naming.md) 发现 [错误处理最佳做法](custom-functions-errors.md)。

## <a name="see-also"></a>另请参阅

* [手动为自定义函数创建 JSON 元数据](custom-functions-json.md)
* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
