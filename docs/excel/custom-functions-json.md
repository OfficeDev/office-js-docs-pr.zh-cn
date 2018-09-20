# <a name="custom-function-metadata"></a>自定义函数元数据

如果在 Excel 加载项中包括[自定义函数](custom-functions-overview.md)，你必须托管一个 JSON 文件，其中包含有关函数的元数据（此外，还要托管包含函数的 JavaScript文件，以及充当 JavaScript 文件父项的无用户界面的 HTML 文件）。 本文使用示例描述了 JSON 文件的格式。

[此处](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)提供了一个完整的示例 JSON文件。

## <a name="functions-array"></a>函数数组

元数据是包含单一 `functions` 属性的 JSON 对象，其值是一个对象数组。 其中的每个对象都代表一个自定义函数。 下表包含其属性：

|  属性  |  数据类型  |  是否必需？  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  字符串  |  否  |  Excel 用户界面中显示的函数的说明。 例如，“将摄氏度值转换为华氏度”。 |
|  `helpUrl`  |  字符串  |   否  |  用户可在其中获得函数相关帮助的 URL。 （它显示在任务窗格中。）例如，“http://contoso.com/help/convertcelsiustofahrenheit.html”  |
|  `name`  |  字符串  |  是  |  用户选择函数时出现在 Excel 用户界面中的函数名称（以命名空间为前缀）。 它应该与 JavaScript 中定义的函数名称相同。 |
|  `options`  |  对象  |  否  |  配置 Excel 处理函数的方式。 有关详细信息，请参阅[选项对象](#options-object)。 |
|  `parameters`  |  数组  |  是  |  有关函数参数的元数据。 有关详细信息，请参阅[参数数组](#parameters-array)。 |
|  `result`  |  对象  |  是  |  有关函数返回的值的元数据。 有关详细信息，请参阅[结果对象](#result-object)。 |

## <a name="options-object"></a>Options 对象

对象配置 Excel 处理函数的 方式。`options` 下表包含其属性：

|  属性  |  数据类型  |  是否必需？  |  说明  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  布尔值  |  否，默认值为 `false` 。  |  如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。如果您使用此选项，Excel 将使用其他 `caller` 参数调用 JavaScript 函数。（请***不要***在 `parameters` 属性中注册此参数）。在函数的正文中， 必须将处理程序分配给 `caller.onCanceled` 成员。|
|  `stream`  |  布尔值  |  否，默认值为 `false` 。  |  如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。 此选项对于快速变化的数据源（如股票价格）非常有用。 如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。 （请***不要*** 在 `parameters` 属性中注册此参数）。 函数不应存在 `return` 语句。 相反，结果值将作为 `caller.setResult` 回调方法的参数传递。|

## <a name="parameters-array"></a>参数数组

属性是一个对象数组。`parameters` 其中每个对象代表一个参数。 下表包含其属性：

|  属性  |  数据类型  |  是否必需？  |  说明  |
|:-----|:-----|:-----|:-----|
|  `description`  |  字符串  |  否 |  参数的描述。  |
|  `dimensionality`  |  字符串  |  是  |  必须是“标量”（即非数组值）或“矩阵”（即一系列行数组）。  |
|  `name`  |  字符串  |  是  |  参数的名称。 此名称显示在 Excel 的 IntelliSense 中。  |
|  `type`  |  字符串  |  是  |  参数的数据类型。 必须为“布尔值”、“数字”或“字符串”。  |

## <a name="result-object"></a>结果对象

属性提供有关函数返回的值的元数据。`results` 下表包含其属性：

|  属性  |  数据类型  |  是否必需？  |  说明  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  字符串  |  否  |  必须是“标量”（即非数组值）或“矩阵”（即一系列行数组）。  |
|  `type`  |  字符串  |  是  |  参数的数据类型。 必须为“布尔值”、“数字”或“字符串”。  |

## <a name="example"></a>示例

以下 JSON 代码是自定义函数元数据的示例。

```json
{
    "functions": [
        {
            "name": "ADD42", 
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "name": "ADD42ASYNC", 
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "increment",
                    "description": "the number to be added each time",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "stream": true,
                "cancelable": true
            }
        },
        {
            "name": "SECONDHIGHEST", 
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "range",
                    "description": "the input range",
                    "type": "number",
                    "dimensionality": "matrix"
                }
            ]
        }
    ]
}

```

## <a name="see-also"></a>另请参阅
[自定义函数](custom-functions-overview.md)<br>
[有关数组公式的指导和示例](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
