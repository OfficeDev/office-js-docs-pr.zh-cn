---
ms.date: 11/06/2020
description: 本地化您的 Excel 自定义函数。
title: 本地化自定义函数
localization_priority: Normal
ms.openlocfilehash: b393cbb76e4993eb77df8ddbe60247c8af74c580
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071653"
---
# <a name="localize-custom-functions"></a>本地化自定义函数

您可以本地化您的外接程序和自定义函数名称。 若要执行此操作，请在函数的 JSON 文件中提供本地化的函数名称，并在 XML 清单文件中提供区域设置信息。

>[!IMPORTANT]
> 自动生成的元数据不能用于本地化，因此您需要手动更新 JSON 文件。 若要了解如何执行此操作，请参阅 [手动创建自定义函数的 JSON 元数据](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>本地化函数名称

若要本地化自定义函数，请为每种语言创建一个新的 JSON 元数据文件。 在每个语言 JSON 文件中 `name` ， `description` 在目标语言中创建和属性。 英语的默认文件命名为 **"functions.js"** 。 对每个其他 JSON 文件使用文件名中的区域设置，如中的 **functions-de.js** ，以帮助识别这些文件。

`name`并 `description` 将显示在 Excel 中并进行本地化。 但是，不会对 `id` 每个函数的进行本地化。 `id`属性是 Excel 将函数标识为唯一的，不应在设置后更改。

下面的 JSON 演示如何定义 `id` 属性 "乘法" 的函数。 `name`函数的和 `description` 属性本地化为德语。 每个 `name` 参数 `description` 也本地化为德语。

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

将以前的 JSON 与以下 JSON 进行比较，以获取英语。

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

## <a name="localize-your-add-in"></a>本地化你的外接程序

为每种语言创建一个 JSON 文件后，使用指定每个 JSON 元数据文件的 URL 的每个区域设置的替代值更新 XML 清单文件。 下面的清单 XML 显示 `en-us` (德国) 的覆盖 JSON 文件 URL 的默认区域设置 `de-de` 。 文件 **上的functions-de.js** 包含本地化的德语函数名称和 id。

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

有关本地化外接程序的过程的详细信息，请参阅 [Office 外接程序的本地化](../develop/localization.md#control-localization-from-the-manifest)。

## <a name="next-steps"></a>后续步骤
了解 [自定义函数的命名约定](custom-functions-naming.md) 或发现 [错误处理最佳实践](custom-functions-errors.md)。

## <a name="see-also"></a>另请参阅

* [手动创建自定义函数的 JSON 元数据](custom-functions-json.md)
* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
