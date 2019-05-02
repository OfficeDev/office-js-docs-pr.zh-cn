---
ms.date: 04/30/2019
description: 本地化您的 Excel 自定义函数。
title: 本地化自定义函数 (预览)
localization_priority: Normal
ms.openlocfilehash: 1c7fba297996c8cf050eb23b34823debf87b4e88
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527302"
---
# <a name="localize-custom-functions"></a>本地化自定义函数

若要使自定义函数在世界范围内正常工作, 请将它们本地化为不同的语言。 若要本地化自定义函数, 需要在函数的 JSON 文件中提供本地化的函数名称, 并在 XML 清单文件中提供区域设置信息。 自动生成的元数据不能用于本地化, 因此您需要手动更新 JSON 文件。

## <a name="localize-function-names"></a>本地化函数名称

若要本地化自定义函数, 请为每种语言创建一个新的 JSON 元数据文件。 在每个语言 JSON 文件中`name` , `description`在目标语言中创建和属性。 英语的默认文件命名为**函数 json**。 建议您在文件名中为每个附加的 JSON 文件 (如**函数-。 JSON** ) 使用区域设置来帮助识别这些文件。 

`name`并`description`将显示在 Excel 中并进行本地化。 但是, 每`id`个函数的不本地化。 `id`属性是 Excel 将函数标识为唯一的, 并且在设置后不应更改。

下面的 JSON 演示如何定义`id`属性 "乘法" 的函数。 函数`name`的`description`和属性本地化为德语。 每个`name`参数`description`也本地化为德语。

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

将以前的 JSON 与以下 JSON 进行比较, 以获取英语。

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

为每种语言创建一个 JSON 文件后, 需要使用每个指定的 JSON 元数据文件的 URL 的区域设置来更新 XML 清单文件的替代值。 下面的清单 XML 显示了一个`en-us`默认区域设置, 其中包含`de-de` (德国) 的覆盖 JSON 文件 URL。 **函数-。 json**文件包含本地化的德语函数名称和 id。

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


有关本地化外接程序的过程的详细信息, 请参阅[Office 外接程序的本地化](../develop/localization.md#control-localization-from-the-manifest)。

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [自定义函数元数据](custom-functions-json.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数更改日志](custom-functions-changelog.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
