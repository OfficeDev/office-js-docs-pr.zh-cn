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
# <a name="localize-custom-functions"></a>本地化自定义函数

您可以本地化您的外接程序和自定义函数名称。 若要执行此操作，请在函数的 JSON 文件中提供本地化的函数名称，并在 XML 清单文件中提供区域设置信息。

>[!IMPORTANT]
> 自动生成的元数据不能用于本地化，因此您需要手动更新 JSON 文件。 若要了解如何执行此操作，请参阅[Excel 中的自定义函数的元数据](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>本地化函数名称

若要本地化自定义函数，请为每种语言创建一个新的 JSON 元数据文件。 在每个语言 JSON 文件中 `name` ， `description` 在目标语言中创建和属性。 英语的默认文件命名为**函数 json**。 对每个其他 JSON 文件使用文件名中的区域设置（如**函数-** 为帮助识别它们）。

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

为每种语言创建一个 JSON 文件后，使用指定每个 JSON 元数据文件的 URL 的每个区域设置的替代值更新 XML 清单文件。 下面的清单 XML 显示了一个默认 `en-us` 区域设置，其中包含 `de-de` （德国）的覆盖 JSON 文件 URL。 **函数-. json**文件包含本地化的德语函数名称和 id。

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

有关本地化外接程序的过程的详细信息，请参阅[Office 外接程序的本地化](../develop/localization.md#control-localization-from-the-manifest)。

## <a name="next-steps"></a>后续步骤
了解[自定义函数的命名约定](custom-functions-naming.md)或发现[错误处理最佳实践](custom-functions-errors.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
