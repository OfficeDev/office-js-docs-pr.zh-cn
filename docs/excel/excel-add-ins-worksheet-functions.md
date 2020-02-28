---
title: 使用 Excel JavaScript API 调用内置 Excel 工作表函数
description: ''
ms.date: 12/19/2019
localization_priority: Normal
ms.openlocfilehash: c5b725f09c4bd6be8d6061f08fe7fbf84ff30762
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325148"
---
# <a name="call-built-in-excel-worksheet-functions"></a>调用内置 Excel 工作表函数

本文介绍了如何使用 Excel JavaScript API 调用内置 Excel 工作表函数（如 `VLOOKUP` 和 `SUM`）。 其中还收录了可以使用 Excel JavaScript API 调用的内置 Excel 工作表函数的完整列表。

> [!NOTE]
> 若要了解如何使用 Excel JavaScript API 在 Excel 中创建*自定义函数*，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。

## <a name="calling-a-worksheet-function"></a>创建工作表函数

下面的代码片段展示了如何调用工作表函数，其中 `sampleFunction()` 是占位符，应将它替换为要调用的函数名称和函数需要使用的输入参数。 工作`value`表函数返回`FunctionResult`的对象的属性包含指定函数的结果。 如以下示例所示，您`load`必须`value`具有该`FunctionResult`对象的属性，然后才能阅读该对象。 在此示例中，函数结果被直接写入控制台。

```js
var functionResult = context.workbook.functions.sampleFunction();
functionResult.load('value');
return context.sync()
    .then(function () {
        console.log('Result of the function: ' + functionResult.value);
    });
```

> [!TIP]
> 有关可以使用 Excel JavaScript API 调用的函数列表，请参阅本文的[支持的工作表函数](#supported-worksheet-functions)部分。

## <a name="sample-data"></a>示例数据

下图展示了 Excel 工作表中的表格，其中包含三个月内各种工具的销售数据。 表格中的每个数字均表示具体工具在特定月份中的销售件数。 接下来的两个示例展示了如何向此类数据应用内置工作表函数。

![锤子、扳手和锯子的 11 月、12 月和 1 月销售数据的 Excel 屏幕截图](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>示例 1：单函数

下面的代码示例向前述示例数据应用 `VLOOKUP` 函数，以确定 11 月售出的扳手数。

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="example-2-nested-functions"></a>示例 2：嵌套函数

下面的代码示例向前述示例数据应用 `VLOOKUP` 函数，以分别确定 11 月和 12 月售出的扳手数。然后，应用 `SUM` 函数，以计算这两个月售出的扳手总数。

如此示例所示，如果一个或多个函数调用嵌套在另一个函数调用中，只需对随后要读取的最终结果（在此示例中为 `load`）执行 `sumOfTwoLookups` 操作即可。 系统会计算所有中间结果（在此示例中，为每个 `VLOOKUP` 函数的结果），并根据这些中间结果计算最终结果。

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false),
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="supported-worksheet-functions"></a>支持的工作表函数

可以使用 Excel JavaScript API 调用以下内置 Excel 工作表函数。

| 功能 | 说明 |
|:---------------|:-----------|
| <a href="https://support.office.com/article/ABS-function-3420200f-5628-4e8c-99da-c99d7c87713c" target="_blank">ABS 函数</a> | 返回数字的绝对值 |
| <a href="https://support.office.com/article/ACCRINT-function-fe45d089-6722-4fb3-9379-e1f911d8dc74" target="_blank">ACCRINT 函数</a> | 返回定期支付利息的债券的应计利息 |
| <a href="https://support.office.com/article/ACCRINTM-function-f62f01f9-5754-4cc4-805b-0e70199328a7" target="_blank">ACCRINTM 函数</a> | 返回在到期日支付利息的债券的应计利息 |
| <a href="https://support.office.com/article/ACOS-function-cb73173f-d089-4582-afa1-76e5524b5d5b" target="_blank">ACOS 函数</a> | 返回一个数的反余弦值 |
| <a href="https://support.office.com/article/ACOSH-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe" target="_blank">ACOSH 函数</a> | 返回一个数的反双曲余弦值 |
| <a href="https://support.office.com/article/ACOT-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905" target="_blank">ACOT 函数</a> | 返回一个数的反余切值 |
| <a href="https://support.office.com/article/ACOTH-function-cc49480f-f684-4171-9fc5-73e4e852300f" target="_blank">ACOTH 函数</a> | 返回一个数的双曲反余切值 |
| <a href="https://support.office.com/article/AMORDEGRC-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51" target="_blank">AMORDEGRC 函数</a> | 通过使用折旧系数，返回每个会计期间的折旧值 |
| <a href="https://support.office.com/article/AMORLINC-function-7d417b45-f7f5-4dba-a0a5-3451a81079a8" target="_blank">AMORLINC 函数</a> | 返回每个会计期间的折旧值 |
| <a href="https://support.office.com/article/AND-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9" target="_blank">AND 函数</a> | 如果所有参数都为 true，返回 `TRUE` |
| <a href="https://support.office.com/article/ARABIC-function-9a8da418-c17b-4ef9-a657-9370a30a674f" target="_blank">ARABIC 函数</a> | 将罗马数字转换为阿拉伯数字 |
| <a href="https://support.office.com/article/AREAS-function-8392ba32-7a41-43b3-96b0-3695d2ec6152" target="_blank">AREAS 函数</a> | 返回引用中包含的区域个数 |
| <a href="https://support.office.com/article/ASC-function-0b6abf1c-c663-4004-a964-ebc00b723266" target="_blank">ASC 函数</a> | 将字符串中的全角（双字节）英文字母或片假名更改为半角（单字节）字符 |
| <a href="https://support.office.com/article/ASIN-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347" target="_blank">ASIN 函数</a> | 返回一个数的反正弦值 |
| <a href="https://support.office.com/article/ASINH-function-4e00475a-067a-43cf-926a-765b0249717c" target="_blank">ASINH 函数</a> | 返回一个数的反双曲正弦值 |
| <a href="https://support.office.com/article/ATAN-function-50746fa8-630a-406b-81d0-4a2aed395543" target="_blank">ATAN 函数</a> | 返回一个数的反正切值 |
| <a href="https://support.office.com/article/ATAN2-function-c04592ab-b9e3-4908-b428-c96b3a565033" target="_blank">ATAN2 函数</a> | 返回从 x 坐标和 y 坐标的反正切值 |
| <a href="https://support.office.com/article/ATANH-function-3cd65768-0de7-4f1d-b312-d01c8c930d90" target="_blank">ATANH 函数</a> | 返回某一数字的反双曲正切值 |
| <a href="https://support.office.com/article/AVEDEV-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639" target="_blank">AVEDEV 函数</a> | 返回一组数据点到其算术平均值的绝对偏差的平均值 |
| <a href="https://support.office.com/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6" target="_blank">AVERAGE 函数</a> | 返回其参数的平均值 |
| <a href="https://support.office.com/article/AVERAGEA-function-f5f84098-d453-4f4c-bbba-3d2c66356091" target="_blank">AVERAGEA 函数</a> | 返回其参数的平均值，包括数字、文本和逻辑值 |
| <a href="https://support.office.com/article/AVERAGEIF-function-faec8e2e-0dec-4308-af69-f5576d8ac642" target="_blank">AVERAGEIF 函数</a> | 返回区域内满足给定条件的所有单元格的平均值（算术平均值） |
| <a href="https://support.office.com/article/AVERAGEIFS-function-48910c45-1fc0-4389-a028-f7c5c3001690" target="_blank">AVERAGEIFS 函数</a> | 返回满足多个条件的所有单元格的平均值（算术平均值） |
| <a href="https://support.office.com/article/BAHTTEXT-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c" target="_blank">BAHTTEXT 函数</a> | 使用 ß（铢）货币格式将数字转换为文本 |
| <a href="https://support.office.com/article/BASE-function-2ef61411-aee9-4f29-a811-1c42456c6342" target="_blank">BASE 函数</a> | 将数字转换成具有给定基数的文本表示形式 |
| <a href="https://support.office.com/article/BESSELI-function-8d33855c-9a8d-444b-98e0-852267b1c0df" target="_blank">BESSELI 函数</a> | 返回修正的贝塞耳函数 In(x) |
| <a href="https://support.office.com/article/BESSELJ-function-839cb181-48de-408b-9d80-bd02982d94f7" target="_blank">BESSELJ 函数</a> | 返回贝塞耳函数 Jn(x) |
| <a href="https://support.office.com/article/BESSELK-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70" target="_blank">BESSELK 函数</a> | 返回修正的贝塞耳函数 Kn(x) |
| <a href="https://support.office.com/article/BESSELY-function-f3a356b3-da89-42c3-8974-2da54d6353a2" target="_blank">BESSELY 函数</a> | 返回贝赛耳函数 Yn(x) |
| <a href="https://support.office.com/article/BETADIST-function-11188c9c-780a-42c7-ba43-9ecb5a878d31" target="_blank">BETA.DIST 函数</a> | 返回 beta 累积分布函数 |
| <a href="https://support.office.com/article/BETAINV-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb" target="_blank">BETA.INV 函数</a> | 返回指定的 beta 分布累积分布函数的反函数 |
| <a href="https://support.office.com/article/BIN2DEC-function-63905b57-b3a0-453d-99f4-647bb519cd6c" target="_blank">BIN2DEC 函数</a> | 将二进制数转换为十进制 |
| <a href="https://support.office.com/article/BIN2HEX-function-0375e507-f5e5-4077-9af8-28d84f9f41cc" target="_blank">BIN2HEX 函数</a> | 将二进制数转换为十六进制 |
| <a href="https://support.office.com/article/BIN2OCT-function-0a4e01ba-ac8d-4158-9b29-16c25c4c23fd" target="_blank">BIN2OCT 函数</a> | 将二进制数转换为八进制 |
| <a href="https://support.office.com/article/BINOMDIST-function-c5ae37b6-f39c-4be2-94c2-509a1480770c" target="_blank">BINOM.DIST 函数</a> | 返回一元二项式分布的概率 |
| <a href="https://support.office.com/article/BINOMDISTRANGE-function-17331329-74c7-4053-bb4c-6653a7421595" target="_blank">BINOM.DIST.RANGE 函数</a> | 返回使用二项式分布的试验结果的概率 |
| <a href="https://support.office.com/article/BINOMINV-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9" target="_blank">BINOM.INV 函数</a> | 返回一个数值，它是使得累积二项式分布的函数值小于或等于临界值的最小整数 |
| <a href="https://support.office.com/article/BITAND-function-8a2be3d7-91c3-4b48-9517-64548008563a" target="_blank">BITAND 函数</a> | 返回两个数字的“按位与” |
| <a href="https://support.office.com/article/BITLSHIFT-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c" target="_blank">BITLSHIFT 函数</a> | 返回按照 shift_amount 位数左移后得到的数值 |
| <a href="https://support.office.com/article/BITOR-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2" target="_blank">BITOR 函数</a> | 返回 2 个数字的按位“或” |
| <a href="https://support.office.com/article/BITRSHIFT-function-274d6996-f42c-4743-abdb-4ff95351222c" target="_blank">BITRSHIFT 函数</a> | 返回按照 shift_amount 位数右移后得到的数值 |
| <a href="https://support.office.com/article/BITXOR-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4" target="_blank">BITXOR Function</a> | 返回两个数字的按位“异或”值 |
| <a href="https://support.office.com/article/CEILINGMATH-function-80f95d2f-b499-4eee-9f16-f795a8e306c8" target="_blank">向上.数学、ECMA_CEILING 函数</a> | 将数值向上舍入为最接近的整数或最接近的基数的倍数 |
| <a href="https://support.office.com/article/CEILINGPRECISE-function-f366a774-527a-4c92-ba49-af0a196e66cb" target="_blank">CEILING.PRECISE 函数</a> | 将数值四舍五入到最接近的整数或最接近的基数的倍数。不论数字是否带有符号，都将数字向上舍入。 |
| <a href="https://support.office.com/article/CHAR-function-bbd249c8-b36e-4a91-8017-1c133f9b837a" target="_blank">CHAR 函数</a> | 返回由代码数字指定的字符 |
| <a href="https://support.office.com/article/CHISQDIST-function-8486b05e-5c05-4942-a9ea-f6b341518732" target="_blank">CHISQ.DIST 函数</a> | 返回累积 beta 分布的概率密度函数 |
| <a href="https://support.office.com/article/CHISQDISTRT-function-dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2" target="_blank">CHISQ.DIST.RT 函数</a> | 返回 χ2 分布的收尾概率 |
| <a href="https://support.office.com/article/CHISQINV-function-400db556-62b3-472d-80b3-254723e7092f" target="_blank">CHISQ.INV 函数</a> | 返回累积 beta 分布的概率密度函数 |
| <a href="https://support.office.com/article/CHISQINVRT-function-435b5ed8-98d5-4da6-823f-293e2cbc94fe" target="_blank">CHISQ.INV.RT 函数</a> | 返回 χ2 分布的收尾概率的反函数 |
| <a href="https://support.office.com/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc" target="_blank">CHOOSE 函数</a> | 从值列表中选择一个值 |
| <a href="https://support.office.com/article/CLEAN-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41" target="_blank">CLEAN 函数</a> | 删除文本中的所有非打印字符 |
| <a href="https://support.office.com/article/CODE-function-c32b692b-2ed0-4a04-bdd9-75640144b928" target="_blank">CODE 函数</a> | 返回文本字符串中第一个字符的数字代码 |
| <a href="https://support.office.com/article/COLUMNS-function-4e8e7b4e-e603-43e8-b177-956088fa48ca" target="_blank">COLUMNS 函数</a> | 返回引用中的列数 |
| <a href="https://support.office.com/article/COMBIN-function-12a3f276-0a21-423a-8de6-06990aaf638a" target="_blank">COMBIN 函数</a> | 返回给定数目对象的组合数 |
| <a href="https://support.office.com/article/COMBINA-function-efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d" target="_blank">COMBINA 函数</a> | 返回给定项数的组合数（包含重复项） |
| <a href="https://support.office.com/article/COMPLEX-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128" target="_blank">COMPLEX 函数</a> | 将实部系数和虚部系数转换为复数 |
| <a href="https://support.office.com/article/CONCATENATE-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d" target="_blank">CONCATENATE 函数</a> | 将几个文本项合并为一个文本项 |
| <a href="https://support.office.com/article/CONFIDENCENORM-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4" target="_blank">CONFIDENCE.NORM 函数</a> | 返回总体平均数的置信区间 |
| <a href="https://support.office.com/article/CONFIDENCET-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53" target="_blank">CONFIDENCE.T 函数</a> | 使用学生 t 分布返回总体平均数的置信区间 |
| <a href="https://support.office.com/article/CONVERT-function-d785bef1-808e-4aac-bdcd-666c810f9af2" target="_blank">CONVERT 函数</a> | 将数字从一种度量体系转换为另一种度量体系 |
| <a href="https://support.office.com/article/COS-function-0fb808a5-95d6-4553-8148-22aebdce5f05" target="_blank">COS 函数</a> | 返回一个数的余弦值 |
| <a href="https://support.office.com/article/COSH-function-e460d426-c471-43e8-9540-a57ff3b70555" target="_blank">COSH 函数</a> | 返回一个数字的双曲余弦值 |
| <a href="https://support.office.com/article/COT-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a" target="_blank">COT 函数</a> | 返回一个角度的余切值 |
| <a href="https://support.office.com/article/COTH-function-2e0b4cb6-0ba0-403e-aed4-deaa71b49df5" target="_blank">COTH 函数</a> | 返回一个数字的双曲余切值 |
| <a href="https://support.office.com/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c" target="_blank">COUNT 函数</a> | 计算参数表中的数字个数 |
| <a href="https://support.office.com/article/COUNTA-function-7dc98875-d5c1-46f1-9a82-53f3219e2509" target="_blank">COUNTA 函数</a> | 计算参数列表中值的数量 |
| <a href="https://support.office.com/article/COUNTBLANK-function-6a92d772-675c-4bee-b346-24af6bd3ac22" target="_blank">COUNTBLANK 函数</a> | 计算在一定范围内的空单元格数量 |
| <a href="https://support.office.com/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34" target="_blank">COUNTIF 函数</a> | 计算某个区域中满足给定条件的单元格数目 |
| <a href="https://support.office.com/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842" target="_blank">COUNTIFS 函数</a> | 计算某个区域中满足多个条件的单元格数目 |
| <a href="https://support.office.com/article/COUPDAYBS-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872" target="_blank">COUPDAYBS 函数</a> | 返回从票息期开始到结算日之间的天数 |
| <a href="https://support.office.com/article/COUPDAYS-function-cc64380b-315b-4e7b-950c-b30b0a76f671" target="_blank">COUPDAYS 函数</a> | 返回包含结算日的票息期的天数 |
| <a href="https://support.office.com/article/COUPDAYSNC-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547" target="_blank">COUPDAYSNC 函数</a> | 返回从结算日到下一票息支付日之间的天数 |
| <a href="https://support.office.com/article/COUPNCD-function-fd962fef-506b-4d9d-8590-16df5393691f" target="_blank">COUPNCD 函数</a> | 返回结算日后的下一票息支付日 |
| <a href="https://support.office.com/article/COUPNUM-function-a90af57b-de53-4969-9c99-dd6139db2522" target="_blank">COUPNUM 函数</a> | 返回结算日与到期日之间可支付的票息数 |
| <a href="https://support.office.com/article/COUPPCD-function-2eb50473-6ee9-4052-a206-77a9a385d5b3" target="_blank">COUPPCD 函数</a> | 返回结算日前的上一票息支付日 |
| <a href="https://support.office.com/article/CSC-function-07379361-219a-4398-8675-07ddc4f135c1" target="_blank">CSC 函数</a> | 返回一个角度的余割值 |
| <a href="https://support.office.com/article/CSCH-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb" target="_blank">CSCH 函数</a> | 返回一个角度的双曲余割值 |
| <a href="https://support.office.com/article/CUMIPMT-function-61067bb0-9016-427d-b95b-1a752af0e606" target="_blank">CUMIPMT 函数</a> | 返回两个付款期之间为贷款累积支付的利息 |
| <a href="https://support.office.com/article/CUMPRINC-function-94a4516d-bd65-41a1-bc16-053a6af4c04d" target="_blank">CUMPRINC 函数</a> | 返回两个付款期之间为贷款累积支付的本金 |
| <a href="https://support.office.com/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349" target="_blank">DATE 函数</a> | 返回特定日期的序列号 |
| <a href="https://support.office.com/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252" target="_blank">DATEVALUE 函数</a> | 将以文本表达的日期转换为序列号 |
| <a href="https://support.office.com/article/DAVERAGE-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee" target="_blank">DAVERAGE 函数</a> | 返回所选数据库条目的平均值 |
| <a href="https://support.office.com/article/DAY-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101" target="_blank">DAY 函数</a> | 将序列号转换为月份中的某一天 |
| <a href="https://support.office.com/article/DAYS-function-57740535-d549-4395-8728-0f07bff0b9df" target="_blank">DAYS 函数</a> | 返回两个日期间相差的天数 |
| <a href="https://support.office.com/article/DAYS360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a" target="_blank">DAYS360 函数</a> | 按每年 360 天计算两个日期间相差的天数 |
| <a href="https://support.office.com/article/DB-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7" target="_blank">DB 函数</a> | 使用固定余额递减法返回指定周期内某项资产的折旧值 |
| <a href="https://support.office.com/article/DBCS-function-a4025e73-63d2-4958-9423-21a24794c9e5" target="_blank">DBCS 函数</a> | 将字符串中的半角（单字节）英文字母或片假名更改为全角（双字节）字符 |
| <a href="https://support.office.com/article/DCOUNT-function-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1" target="_blank">DCOUNT 函数</a> | 计算数据库中包含数字的单元格数量 |
| <a href="https://support.office.com/article/DCOUNTA-function-00232a6d-5a66-4a01-a25b-c1653fda1244" target="_blank">DCOUNTA 函数</a> | 计算数据库中的非空单元格的数量 |
| <a href="https://support.office.com/article/DDB-function-519a7a37-8772-4c96-85c0-ed2c209717a5" target="_blank">DDB 函数</a> | 使用双倍余额递减法或其他指定方法返回某项资产在指定周期内的折旧值 |
| <a href="https://support.office.com/article/DEC2BIN-function-0f63dd0e-5d1a-42d8-b511-5bf5c6d43838" target="_blank">DEC2BIN 函数</a> | 将十进制数转换为二进制 |
| <a href="https://support.office.com/article/DEC2HEX-function-6344ee8b-b6b5-4c6a-a672-f64666704619" target="_blank">DEC2HEX 函数</a> | 将十进制数转换为十六进制 |
| <a href="https://support.office.com/article/DEC2OCT-function-c9d835ca-20b7-40c4-8a9e-d3be351ce00f" target="_blank">DEC2OCT 函数</a> | 将十进制数转换为八进制 |
| <a href="https://support.office.com/article/DECIMAL-function-ee554665-6176-46ef-82de-0a283658da2e" target="_blank">DECIMAL 函数</a> | 按给定基数将数字的文本表示形式转换成十进制数 |
| <a href="https://support.office.com/article/DEGREES-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1" target="_blank">DEGREES 函数</a> | 将弧度转换为角度 |
| <a href="https://support.office.com/article/DELTA-function-2f763672-c959-4e07-ac33-fe03220ba432" target="_blank">DELTA 函数</a> | 测试两个值是否相等 |
| <a href="https://support.office.com/article/DEVSQ-function-8b739616-8376-4df5-8bd0-cfe0a6caf444" target="_blank">DEVSQ 函数</a> | 返回偏差平方和 |
| <a href="https://support.office.com/article/DGET-function-455568bf-4eef-45f7-90f0-ec250d00892e" target="_blank">DGET 函数</a> | 从数据库中提取符合指定条件的单个记录 |
| <a href="https://support.office.com/article/DISC-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53" target="_blank">DISC 函数</a> | 返回债券的贴现率 |
| <a href="https://support.office.com/article/DMAX-function-f4e8209d-8958-4c3d-a1ee-6351665d41c2" target="_blank">DMAX 函数</a> | 返回所选数据库条目中的最大值 |
| <a href="https://support.office.com/article/DMIN-function-4ae6f1d9-1f26-40f1-a783-6dc3680192a3" target="_blank">DMIN 函数</a> | 返回所选数据库条目中的最小值 |
| <a href="https://support.office.com/article/DOLLAR-function-a6cd05d9-9740-4ad3-a469-8109d18ff611" target="_blank">美元、USDOLLAR 函数</a> | 使用 $（美元）货币格式将数字转换为文本 |
| <a href="https://support.office.com/article/DOLLARDE-function-db85aab0-1677-428a-9dfd-a38476693427" target="_blank">DOLLARDE 函数</a> | 将以分数表示的货币值转换为以小数表示的货币值 |
| <a href="https://support.office.com/article/DOLLARFR-function-0835d163-3023-4a33-9824-3042c5d4f495" target="_blank">DOLLARFR 函数</a> | 将以小数表示的货币值转换为以分数表示的货币值 |
| <a href="https://support.office.com/article/DPRODUCT-function-4f96b13e-d49c-47a7-b769-22f6d017cb31" target="_blank">DPRODUCT 函数</a> | 将与数据库中的条件匹配的记录的特定字段中的值相乘 |
| <a href="https://support.office.com/article/DSTDEV-function-026b8c73-616d-4b5e-b072-241871c4ab96" target="_blank">DSTDEV 函数</a> | 根据所选数据库条目中的样本估算数据的标准偏差 |
| <a href="https://support.office.com/article/DSTDEVP-function-04b78995-da03-4813-bbd9-d74fd0f5d94b" target="_blank">DSTDEVP 函数</a> | 以数据库选定项作为样本总体，计算数据的标准偏差 |
| <a href="https://support.office.com/article/DSUM-function-53181285-0c4b-4f5a-aaa3-529a322be41b" target="_blank">DSUM 函数</a> | 将数据库中与条件匹配的记录字段列中的数字进行求和 |
| <a href="https://support.office.com/article/DURATION-function-b254ea57-eadc-4602-a86a-c8e369334038" target="_blank">DURATION 函数</a> | 返回定期支付利息的债券的年持续时间 |
| <a href="https://support.office.com/article/DVAR-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1" target="_blank">DVAR 函数</a> | 根据所选数据库条目中的样本估算数据的方差 |
| <a href="https://support.office.com/article/DVARP-function-eb0ba387-9cb7-45c8-81e9-0394912502fc" target="_blank">DVARP 函数</a> | 以数据库选定项作为样本总体，计算数据的总体方差 |
| <a href="https://support.office.com/article/EDATE-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5" target="_blank">EDATE 函数</a> | 返回一串日期，指示起始日期之前/之后的月数 |
| <a href="https://support.office.com/article/EFFECT-function-910d4e4c-79e2-4009-95e6-507e04f11bc4" target="_blank">EFFECT 函数</a> | 返回年有效利率 |
| <a href="https://support.office.com/article/EOMONTH-function-7314ffa1-2bc9-4005-9d66-f49db127d628" target="_blank">EOMONTH 函数</a> | 返回一串日期，表示指定月数之前或之后的月份的最后一天 |
| <a href="https://support.office.com/article/ERF-function-c53c7e7b-5482-4b6c-883e-56df3c9af349" target="_blank">ERF 函数</a> | 返回误差函数 |
| <a href="https://support.office.com/article/ERFPRECISE-function-9a349593-705c-4278-9a98-e4122831a8e0" target="_blank">ERF.PRECISE 函数</a> | 返回误差函数 |
| <a href="https://support.office.com/article/ERFC-function-736e0318-70ba-4e8b-8d08-461fe68b71b3" target="_blank">ERFC 函数</a> | 返回补余误差函数 |
| <a href="https://support.office.com/article/ERFCPRECISE-function-e90e6bab-f45e-45df-b2ac-cd2eb4d4a273" target="_blank">ERFC.PRECISE 函数</a> | 返回在 x 和无穷大之间集成的补余 ERF 函数 |
| <a href="https://support.office.com/article/ERRORTYPE-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa" target="_blank">ERROR.TYPE 函数</a> | 返回对应于一种错误类型的数字 |
| <a href="https://support.office.com/article/EVEN-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9" target="_blank">EVEN 函数</a> | 将数字向上舍入到最近的偶数 |
| <a href="https://support.office.com/article/EXACT-function-d3087698-fc15-4a15-9631-12575cf29926" target="_blank">EXACT 函数</a> | 检查两个文本值是否相同 |
| <a href="https://support.office.com/article/EXP-function-c578f034-2c45-4c37-bc8c-329660a63abe" target="_blank">EXP 函数</a> | 返回 e 的 n 次方 |
| <a href="https://support.office.com/article/EXPONDIST-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e" target="_blank">EXPON.DIST 函数</a> | 返回指数分布 |
| <a href="https://support.office.com/article/FDIST-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d" target="_blank">F.DIST 函数</a> | 返回 F 概率分布 |
| <a href="https://support.office.com/article/FDISTRT-function-d74cbb00-6017-4ac9-b7d7-6049badc0520" target="_blank">F.DIST.RT 函数</a> | 返回 F 概率分布 |
| <a href="https://support.office.com/article/FINV-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe" target="_blank">F.INV 函数</a> | 返回 F 概率分布的逆函数值 |
| <a href="https://support.office.com/article/FINVRT-function-d371aa8f-b0b1-40ef-9cc2-496f0693ac00" target="_blank">F.INV.RT 函数</a> | 返回 F 概率分布的逆函数值 |
| <a href="https://support.office.com/article/FACT-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3" target="_blank">FACT 函数</a> | 返回某数的阶乘 |
| <a href="https://support.office.com/article/FACTDOUBLE-function-e67697ac-d214-48eb-b7b7-cce2589ecac8" target="_blank">FACTDOUBLE 函数</a> | 返回数字的双阶乘 |
| <a href="https://support.office.com/article/FALSE-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904" target="_blank">FALSE 函数</a> | 返回逻辑值 `FALSE` |
| <a href="https://support.office.com/article/FIND-FINDB-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628" target="_blank">FIND、FINDB 函数</a> | 在一个文本值中查找另一个（区分大小写） |
| <a href="https://support.office.com/article/FISHER-function-d656523c-5076-4f95-b87b-7741bf236c69" target="_blank">FISHER 函数</a> | 返回 Fisher 变换值 |
| <a href="https://support.office.com/article/FISHERINV-function-62504b39-415a-4284-a285-19c8e82f86bb" target="_blank">FISHERINV 函数</a> | 返回 Fisher 逆变换值 |
| <a href="https://support.office.com/article/FIXED-function-ffd5723c-324c-45e9-8b96-e41be2a8274a" target="_blank">FIXED 函数</a> | 将数字格式化为具有固定数量的小数的文本 |
| <a href="https://support.office.com/article/FLOORMATH-function-c302b599-fbdb-4177-ba19-2c2b1249a2f5" target="_blank">FLOOR.MATH 函数</a> | 将数字向下舍入为最接近的整数或最接近的基数的倍数 |
| <a href="https://support.office.com/article/FLOORPRECISE-function-f769b468-1452-4617-8dc3-02f842a0702e" target="_blank">FLOOR.PRECISE 函数</a> | 将数字向下舍入为最接近的整数或最接近的基数的倍数。不论数字是否带有符号，都将数字向下舍入。 |
| <a href="https://support.office.com/article/FV-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3" target="_blank">FV 函数</a> | 返回一项投资的未来值 |
| <a href="https://support.office.com/article/FVSCHEDULE-function-bec29522-bd87-4082-bab9-a241f3fb251d" target="_blank">FVSCHEDULE 函数</a> | 返回在应用一系列复利后，初始本金的终值 |
| <a href="https://support.office.com/article/GAMMA-function-ce1702b1-cf55-471d-8307-f83be0fc5297" target="_blank">GAMMA 函数</a> | 返回 Gamma 函数值 |
| <a href="https://support.office.com/article/GAMMADIST-function-9b6f1538-d11c-4d5f-8966-21f6a2201def" target="_blank">GAMMA.DIST 函数</a> | 返回 γ 分布 |
| <a href="https://support.office.com/article/GAMMAINV-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18" target="_blank">GAMMA.INV 函数</a> | 返回 γ 累积分布的反函数 |
| <a href="https://support.office.com/article/GAMMALN-function-b838c48b-c65f-484f-9e1d-141c55470eb9" target="_blank">GAMMALN 函数</a> | 返回 γ 函数的自然对数 Γ(x) |
| <a href="https://support.office.com/article/GAMMALNPRECISE-function-5cdfe601-4e1e-4189-9d74-241ef1caa599" target="_blank">GAMMALN.PRECISE 函数</a> | 返回 γ 函数的自然对数 Γ(x) |
| <a href="https://support.office.com/article/GAUSS-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33" target="_blank">GAUSS 函数</a> | 返回比标准正态累积分布小 0.5 的值 |
| <a href="https://support.office.com/article/GCD-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a" target="_blank">GCD 函数</a> | 返回最大公约数 |
| <a href="https://support.office.com/article/GEOMEAN-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5" target="_blank">GEOMEAN 函数</a> | 返回几何平均数 |
| <a href="https://support.office.com/article/GESTEP-function-f37e7d2a-41da-4129-be95-640883fca9df" target="_blank">GESTEP 函数</a> | 测试某个数字是否大于阈值 |
| <a href="https://support.office.com/article/HARMEAN-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6" target="_blank">HARMEAN 函数</a> | 返回调和平均值 |
| <a href="https://support.office.com/article/HEX2BIN-function-a13aafaa-5737-4920-8424-643e581828c1" target="_blank">HEX2BIN 函数</a> | 将十六进制数转换为二进制 |
| <a href="https://support.office.com/article/HEX2DEC-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e" target="_blank">HEX2DEC 函数</a> | 将十六进制数转换为十进制 |
| <a href="https://support.office.com/article/HEX2OCT-function-54d52808-5d19-4bd0-8a63-1096a5d11912" target="_blank">HEX2OCT 函数</a> | 将十六进制数转换为八进制 |
| <a href="https://support.office.com/article/HLOOKUP-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f" target="_blank">HLOOKUP 函数</a> | 在数组的顶行中查找并返回指定单元格的值 |
| <a href="https://support.office.com/article/HOUR-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7" target="_blank">HOUR 函数</a> | 将序列号转换为小时 |
| <a href="https://support.office.com/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f" target="_blank">HYPERLINK 函数</a> | 创建一个快捷方式或链接，以便打开一个存储在网络服务器、内部网或 Internet 上的文档 |
| <a href="https://support.office.com/article/HYPGEOMDIST-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf" target="_blank">HYPGEOM.DIST 函数</a> | 返回超几何分布 |
| <a href="https://support.office.com/article/IF-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2" target="_blank">IF 函数</a> | 指定要执行的逻辑测试 |
| <a href="https://support.office.com/article/IMABS-function-b31e73c6-d90c-4062-90bc-8eb351d765a1" target="_blank">IMABS 函数</a> | 返回复数的绝对值（模数） |
| <a href="https://support.office.com/article/IMAGINARY-function-dd5952fd-473d-44d9-95a1-9a17b23e428a" target="_blank">IMAGINARY 函数</a> | 返回复数的虚部系数 |
| <a href="https://support.office.com/article/IMARGUMENT-function-eed37ec1-23b3-4f59-b9f3-d340358a034a" target="_blank">IMARGUMENT 函数</a> | 返回以弧度表示的角 - 参数 θ |
| <a href="https://support.office.com/article/IMCONJUGATE-function-2e2fc1ea-f32b-4f9b-9de6-233853bafd42" target="_blank">IMCONJUGATE 函数</a> | 返回复数的共轭复数 |
| <a href="https://support.office.com/article/IMCOS-function-dad75277-f592-4a6b-ad6c-be93a808a53c" target="_blank">IMCOS 函数</a> | 返回复数的余弦值 |
| <a href="https://support.office.com/article/IMCOSH-function-053e4ddb-4122-458b-be9a-457c405e90ff" target="_blank">IMCOSH 函数</a> | 返回复数的双曲余弦值 |
| <a href="https://support.office.com/article/IMCOT-function-dc6a3607-d26a-4d06-8b41-8931da36442c" target="_blank">IMCOT 函数</a> | 返回复数的余切值 |
| <a href="https://support.office.com/article/IMCSC-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323" target="_blank">IMCSC 函数</a> | 返回复数的余割值 |
| <a href="https://support.office.com/article/IMCSCH-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9" target="_blank">IMCSCH 函数</a> | 返回复数的双曲余割值 |
| <a href="https://support.office.com/article/IMDIV-function-a505aff7-af8a-4451-8142-77ec3d74d83f" target="_blank">IMDIV 函数</a> | 返回两个复数之商 |
| <a href="https://support.office.com/article/IMEXP-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f" target="_blank">IMEXP 函数</a> | 返回复数的指数值 |
| <a href="https://support.office.com/article/IMLN-function-32b98bcf-8b81-437c-a636-6fb3aad509d8" target="_blank">IMLN 函数</a> | 返回复数的自然对数 |
| <a href="https://support.office.com/article/IMLOG10-function-58200fca-e2a2-4271-8a98-ccd4360213a5" target="_blank">IMLOG10 函数</a> | 返回以 10 为底的复数的对数 |
| <a href="https://support.office.com/article/IMLOG2-function-152e13b4-bc79-486c-a243-e6a676878c51" target="_blank">IMLOG2 函数</a> | 返回以 2 为底的复数的对数 |
| <a href="https://support.office.com/article/IMPOWER-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39" target="_blank">IMPOWER 函数</a> | 返回复数的整数幂 |
| <a href="https://support.office.com/article/IMPRODUCT-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba" target="_blank">IMPRODUCT 函数</a> | 返回从 2 到 255 个复数的乘积 |
| <a href="https://support.office.com/article/IMREAL-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366" target="_blank">IMREAL 函数</a> | 返回复数的实部系数 |
| <a href="https://support.office.com/article/IMSEC-function-6df11132-4411-4df4-a3dc-1f17372459e0" target="_blank">IMSEC 函数</a>IMSEC 函数 | 返回复数的正割值 |
| <a href="https://support.office.com/article/IMSECH-function-f250304f-788b-4505-954e-eb01fa50903b" target="_blank">IMSECH 函数</a> | 返回复数的双曲正割值 |
| <a href="https://support.office.com/article/IMSIN-function-1ab02a39-a721-48de-82ef-f52bf37859f6" target="_blank">IMSIN 函数</a> | 返回复数的正弦值 |
| <a href="https://support.office.com/article/IMSINH-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d" target="_blank">IMSINH 函数</a> | 返回复数的双曲正弦值 |
| <a href="https://support.office.com/article/IMSQRT-function-e1753f80-ba11-4664-a10e-e17368396b70" target="_blank">IMSQRT 函数</a> | 返回复数的平方根 |
| <a href="https://support.office.com/article/IMSUB-function-2e404b4d-4935-4e85-9f52-cb08b9a45054" target="_blank">IMSUB 函数</a> | 返回两个复数的差值 |
| <a href="https://support.office.com/article/IMSUM-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f" target="_blank">IMSUM 函数</a> | 返回复数的和 |
| <a href="https://support.office.com/article/IMTAN-function-8478f45d-610a-43cf-8544-9fc0b553a132" target="_blank">IMTAN 函数</a> | 返回复数的正切值 |
| <a href="https://support.office.com/article/INT-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef" target="_blank">INT 函数</a> | 将数值向下舍入到最接近的整数 |
| <a href="https://support.office.com/article/INTRATE-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f" target="_blank">INTRATE 函数</a> | 返回完全投资型债券的利率 |
| <a href="https://support.office.com/article/IPMT-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f" target="_blank">IPmt 函数</a> | 返回给定期间内投资所支付的利息 |
| <a href="https://support.office.com/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc" target="_blank">IRR 函数</a> | 返回一系列现金流的内部收益率 |
| <a href="https://support.office.com/article/ISERR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISERR 函数</a> | 如果值是除 #N/A 之外的错误值，返回 `TRUE` |
| <a href="https://support.office.com/article/ISERROR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISERROR 函数</a> | 如果值是任何错误值，返回 `TRUE` |
| <a href="https://support.office.com/article/ISEVEN-function-aa15929a-d77b-4fbb-92f4-2f479af55356" target="_blank">ISEVEN 函数</a> | 如果值是偶数，返回 `TRUE` |
| <a href="https://support.office.com/article/ISFORMULA-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5" target="_blank">ISFORMULA 函数</a> | 如果存在对包含公式的单元格的引用，返回 `TRUE` |
| <a href="https://support.office.com/article/ISLOGICAL-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISLOGICAL 函数</a> | 如果值是逻辑值，返回 `TRUE` |
| <a href="https://support.office.com/article/ISNA-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISNA 函数</a> | 如果值是 #N/A 错误值，返回 `TRUE` |
| <a href="https://support.office.com/article/ISNONTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISNONTEXT 函数</a> | 如果值不是文本，返回 `TRUE` |
| <a href="https://support.office.com/article/ISNUMBER-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISNUMBER 函数</a> | 如果值是数字，返回 `TRUE` |
| <a href="https://support.office.com/article/ISOCEILING-function-e587bb73-6cc2-4113-b664-ff5b09859a83" target="_blank">ISO.CEILING 函数</a> | 将数字向上舍入到最接近的整数或最接近的基数的倍数 |
| <a href="https://support.office.com/article/ISODD-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISODD 函数</a> | 如果值是奇数，返回 `TRUE` |
| <a href="https://support.office.com/article/ISOWEEKNUM-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e" target="_blank">ISOWEEKNUM 函数</a> | 返回一年中给定日期的 ISO 周数的数目 |
| <a href="https://support.office.com/article/ISPMT-function-fa58adb6-9d39-4ce0-8f43-75399cea56cc" target="_blank">ISPMT 函数</a> | 计算指定的投资期间支付的利息 |
| <a href="https://support.office.com/article/ISREF-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISREF 函数</a> | 如果值是引用，返回 `TRUE` |
| <a href="https://support.office.com/article/ISTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">ISTEXT 函数</a> | 如果值是文本，返回 `TRUE` |
| <a href="https://support.office.com/article/KURT-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab" target="_blank">KURT 函数</a> | 返回一组数据的峰值 |
| <a href="https://support.office.com/article/LARGE-function-3af0af19-1190-42bb-bb8b-01672ec00a64" target="_blank">LARGE 函数</a> | 返回数据集中第 k 个最大值 |
| <a href="https://support.office.com/article/LCM-function-7152b67a-8bb5-4075-ae5c-06ede5563c94" target="_blank">LCM 函数</a> | 返回最小公倍数 |
| <a href="https://support.office.com/article/LEFT-LEFTB-functions-9203d2d2-7960-479b-84c6-1ea52b99640c" target="_blank">LEFT、LEFTB 函数</a> | 返回一个文本值的最左端字符 |
| <a href="https://support.office.com/article/LEN-LENB-functions-29236f94-cedc-429d-affd-b5e33d2c67cb" target="_blank">LEN、LENB 函数</a> | 返回文本字符串中的字符数 |
| <a href="https://support.office.com/article/LN-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f" target="_blank">LN 函数</a> | 返回数值的自然对数 |
| <a href="https://support.office.com/article/LOG-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280" target="_blank">LOG 函数</a> | 返回一个数在指定底下的对数 |
| <a href="https://support.office.com/article/LOG10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211" target="_blank">LOG10 函数</a> | 返回以 10 为底的对数 |
| <a href="https://support.office.com/article/LOGNORMDIST-function-eb60d00b-48a9-4217-be2b-6074aee6b070" target="_blank">LOGNORM.DIST 函数</a> | 返回对数正态分布 |
| <a href="https://support.office.com/article/LOGNORMINV-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600" target="_blank">LOGNORM.INV 函数</a> | 返回对数正态分布的反函数 |
| <a href="https://support.office.com/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb" target="_blank">LOOKUP 函数</a> | 在向量或数组中查找值 |
| <a href="https://support.office.com/article/LOWER-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4" target="_blank">LOWER 函数</a> | 将文本转换为小写 |
| <a href="https://support.office.com/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a" target="_blank">MATCH 函数</a> | 在引用或数组中查找值 |
| <a href="https://support.office.com/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098" target="_blank">MAX 函数</a> | 返回参数列表中的最大值 |
| <a href="https://support.office.com/article/MAXA-function-814bda1e-3840-4bff-9365-2f59ac2ee62d" target="_blank">MAXA 函数</a> | 返回参数列表中的最大值，包括数字、文本和逻辑值 |
| <a href="https://support.office.com/article/MDURATION-function-b3786a69-4f20-469a-94ad-33e5b90a763c" target="_blank">MDURATION 函数</a> | 为假定票面值为 100 元的债券返回麦考利修正持续时间 |
| <a href="https://support.office.com/article/MEDIAN-function-d0916313-4753-414c-8537-ce85bdd967d2" target="_blank">MEDIAN 函数</a> | 返回给定数字的中值 |
| <a href="https://support.office.com/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028" target="_blank">MID、MIDB 函数</a> | 从指定位置开始，返回文本字符串中特定数量的字符。 |
| <a href="https://support.office.com/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152" target="_blank">MIN 函数</a> | 返回参数列表中的最小值 |
| <a href="https://support.office.com/article/MINA-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3" target="_blank">MINA 函数</a> | 返回参数列表中的最小值，包括数字、文本和逻辑值 |
| <a href="https://support.office.com/article/MINUTE-function-af728df0-05c4-4b07-9eed-a84801a60589" target="_blank">MINUTE 函数</a> | 将序列号转换为分钟 |
| <a href="https://support.office.com/article/MIRR-function-b020f038-7492-4fb4-93c1-35c345b53524" target="_blank">MIRR 函数</a> | 返回内部收益率，它的正现金流和负现金流以不同的比率融资 |
| <a href="https://support.office.com/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3" target="_blank">MOD 函数</a> | 返回除法的余数 |
| <a href="https://support.office.com/article/MONTH-function-579a2881-199b-48b2-ab90-ddba0eba86e8" target="_blank">MONTH 函数</a> | 将序列号转换为月 |
| <a href="https://support.office.com/article/MROUND-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427" target="_blank">MROUND 函数</a> | 返回舍入到所需倍数的数值 |
| <a href="https://support.office.com/article/MULTINOMIAL-function-6fa6373c-6533-41a2-a45e-a56db1db1bf6" target="_blank">MULTINOMIAL 函数</a> | 返回一组数字的多项式 |
| <a href="https://support.office.com/article/N-function-a624cad1-3635-4208-b54a-29733d1278c9" target="_blank">N 函数</a> | 返回转换为数字的值 |
| <a href="https://support.office.com/article/NA-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c" target="_blank">NA 函数</a> | 返回错误值 #N/A |
| <a href="https://support.office.com/article/NEGBINOMDIST-function-c8239f89-c2d0-45bd-b6af-172e570f8599" target="_blank">NEGBINOM.DIST 函数</a> | 返回负二项式分布函数值 |
| <a href="https://support.office.com/article/NETWORKDAYS-function-48e717bf-a7a3-495f-969e-5005e3eb18e7" target="_blank">NETWORKDAYS 函数</a> | 返回两个日期之间的完整工作日数 |
| <a href="https://support.office.com/article/NETWORKDAYSINTL-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28" target="_blank">NETWORKDAYS.INTL 函数</a> | 使用能够指示哪些以及有多少天是周末的参数返回两个日期之间的完整工作日数 |
| <a href="https://support.office.com/article/NOMINAL-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b" target="_blank">NOMINAL 函数</a> | 返回年度的单利 |
| <a href="https://support.office.com/article/NORMDIST-function-edb1cc14-a21c-4e53-839d-8082074c9f8d" target="_blank">NORM.DIST 函数</a> | 返回正态分布函数值 |
| <a href="https://support.office.com/article/NORMINV-function-54b30935-fee7-493c-bedb-2278a9db7e13" target="_blank">NORM.INV 函数</a> | 返回正态分布的反函数 |
| <a href="https://support.office.com/article/NORMSDIST-function-1e787282-3832-4520-a9ae-bd2a8d99ba88" target="_blank">NORM.S.DIST 函数</a> | 返回标准正态分布函数值 |
| <a href="https://support.office.com/article/NORMSINV-function-d6d556b4-ab7f-49cd-b526-5a20918452b1" target="_blank">NORM.S.INV 函数</a> | 返回标准正态分布的反函数 |
| <a href="https://support.office.com/article/NOT-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77" target="_blank">NOT 函数</a> | 反转其参数的逻辑 |
| <a href="https://support.office.com/article/NOW-function-3337fd29-145a-4347-b2e6-20c904739c46" target="_blank">NOW 函数</a> | 返回当前日期和时间的序列号 |
| <a href="https://support.office.com/article/NPER-function-240535b5-6653-4d2d-bfcf-b6a38151d815" target="_blank">NPER 函数</a> | 返回一项投资的周期数量 |
| <a href="https://support.office.com/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568" target="_blank">NPV 函数</a> | 基于一系列定期现金流和贴现率返回投资的净现值 |
| <a href="https://support.office.com/article/NUMBERVALUE-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879" target="_blank">NUMBERVALUE 函数</a> | 按独立于区域设置的方式将文本转换为数字 |
| <a href="https://support.office.com/article/OCT2BIN-function-55383471-3c56-4d27-9522-1a8ec646c589" target="_blank">OCT2BIN 函数</a> | 将八进制数转换为二进制 |
| <a href="https://support.office.com/article/OCT2DEC-function-87606014-cb98-44b2-8dbb-e48f8ced1554" target="_blank">OCT2DEC 函数</a> | 将八进制数转换为十进制 |
| <a href="https://support.office.com/article/OCT2HEX-function-912175b4-d497-41b4-a029-221f051b858f" target="_blank">OCT2HEX 函数</a> | 将八进制数转换为十六进制 |
| <a href="https://support.office.com/article/ODD-function-deae64eb-e08a-4c88-8b40-6d0b42575c98" target="_blank">ODD 函数</a> | 将数值向上舍入到最接近的奇数 |
| <a href="https://support.office.com/article/ODDFPRICE-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1" target="_blank">ODDFPRICE 函数</a> | 返回每张票面为 100 元且第一期为奇数的债券的现价 |
| <a href="https://support.office.com/article/ODDFYIELD-function-66bc8b7b-6501-4c93-9ce3-2fd16220fe37" target="_blank">ODDFYIELD 函数</a> | 返回第一期为奇数的债券的收益 |
| <a href="https://support.office.com/article/ODDLPRICE-function-fb657749-d200-4902-afaf-ed5445027fc4" target="_blank">ODDLPRICE 函数</a> | 返回每张票面为 100 元且最后一期为奇数的债券的现价 |
| <a href="https://support.office.com/article/ODDLYIELD-function-c873d088-cf40-435f-8d41-c8232fee9238" target="_blank">ODDLYIELD 函数</a> | 返回最后一期为奇数的债券的收益 |
| <a href="https://support.office.com/article/OR-function-7d17ad14-8700-4281-b308-00b131e22af0" target="_blank">OR 函数</a> | 如果任意参数为 true，返回 `TRUE` |
| <a href="https://support.office.com/article/PDURATION-function-44f33460-5be5-4c90-b857-22308892adaf" target="_blank">PDURATION 函数</a> | 返回投资达到指定的值所需的期数 |
| <a href="https://support.office.com/article/PERCENTILEEXC-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba" target="_blank">PERCENTILE.EXC 函数</a> | 返回数组的 K 百分点值，K 介于 0 与 1 之间，不含 0 与 1 |
| <a href="https://support.office.com/article/PERCENTILEINC-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed" target="_blank">PERCENTILE.INC 函数</a> | 返回数组的 K 百分点值 |
| <a href="https://support.office.com/article/PERCENTRANKEXC-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314" target="_blank">PERCENTRANK.EXC 函数</a> | 返回特定数值在一个数据集中的百分比排名（介于 0 与 1 之间，不含 0 与 1） |
| <a href="https://support.office.com/article/PERCENTRANKINC-function-149592c9-00c0-49ba-86c1-c1f45b80463a" target="_blank">PERCENTRANK.INC 函数</a> | 返回一组数据中的值的百分比排名 |
| <a href="https://support.office.com/article/PERMUT-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3" target="_blank">PERMUT 函数</a> | 返回给定数目对象的排列数 |
| <a href="https://support.office.com/article/PERMUTATIONA-function-6c7d7fdc-d657-44e6-aa19-2857b25cae4e" target="_blank">PERMUTATIONA 函数</a> | 返回从给定元素数目的集合中选取若干（包括重复项）元素的排列数 |
| <a href="https://support.office.com/article/PHI-function-23e49bc6-a8e8-402d-98d3-9ded87f6295c" target="_blank">PHI 函数</a> | 返回标准正态分布的密度函数值 |
| <a href="https://support.office.com/article/PI-function-264199d0-a3ba-46b8-975a-c4a04608989b" target="_blank">PI 函数</a> | 返回 pi 值 |
| <a href="https://support.office.com/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441" target="_blank">PMT 函数</a> | 返回年金的定期支付额 |
| <a href="https://support.office.com/article/POISSONDIST-function-8fe148ff-39a2-46cb-abf3-7772695d9636" target="_blank">POISSON.DIST 函数</a> | 返回泊松分布 |
| <a href="https://support.office.com/article/POWER-function-d3f2908b-56f4-4c3f-895a-07fb519c362a" target="_blank">POWER 函数</a> | 返回某数的乘幂结果 |
| <a href="https://support.office.com/article/PPMT-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b" target="_blank">PPMT 函数</a> | 返回对给定期间内的投资所支付的本金 |
| <a href="https://support.office.com/article/PRICE-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a" target="_blank">PRICE 函数</a> | 返回每张票面为 100 元且定期支付利息的债券的现价 |
| <a href="https://support.office.com/article/PRICEDISC-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3" target="_blank">PRICEDISC 函数</a> | 返回每张票面为 100 元的已贴现债券的现价 |
| <a href="https://support.office.com/article/PRICEMAT-function-52c3b4da-bc7e-476a-989f-a95f675cae77" target="_blank">PRICEMAT 函数</a> | 返回每张票面为 100 元且在到期日支付利息的债券的现价 |
| <a href="https://support.office.com/article/PRODUCT-function-8e6b5b24-90ee-4650-aeec-80982a0512ce" target="_blank">PRODUCT 函数</a> | 将其参数相乘 |
| <a href="https://support.office.com/article/PROPER-function-52a5a283-e8b2-49be-8506-b2887b889f94" target="_blank">PROPER 函数</a> | 使一个文本值的每个词的首字母大写 |
| <a href="https://support.office.com/article/PV-function-23879d31-0e02-4321-be01-da16e8168cbd" target="_blank">PV 函数</a> | 返回一项投资的当前值 |
| <a href="https://support.office.com/article/QUARTILEEXC-function-5a355b7a-840b-4a01-b0f1-f538c2864cad" target="_blank">QUARTILE.EXC 函数</a> | 基于从 0 到 1 之间（不含 0 与 1）的百分点值，返回一组数据的四分位点 |
| <a href="https://support.office.com/article/QUARTILEINC-function-1bbacc80-5075-42f1-aed6-47d735c4819d" target="_blank">QUARTILE.INC 函数</a> | 返回一组数据的四分位点 |
| <a href="https://support.office.com/article/QUOTIENT-function-9f7bf099-2a18-4282-8fa4-65290cc99dee" target="_blank">QUOTIENT 函数</a> | 返回除法结果的整数部分 |
| <a href="https://support.office.com/article/RADIANS-function-ac409508-3d48-45f5-ac02-1497c92de5bf" target="_blank">RADIANS 函数</a> | 将度转换为弧度 |
| <a href="https://support.office.com/article/RAND-function-4cbfa695-8869-4788-8d90-021ea9f5be73" target="_blank">RAND 函数</a> | 返回 0 和 1 之间的一个随机数 |
| <a href="https://support.office.com/article/RANDBETWEEN-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685" target="_blank">RANDBETWEEN 函数</a> | 返回指定数字之间的随机数 |
| <a href="https://support.office.com/article/RANKAVG-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a" target="_blank">RANK.AVG 函数</a> | 返回某数字在一列数字中的排名 |
| <a href="https://support.office.com/article/RANKEQ-function-284858ce-8ef6-450e-b662-26245be04a40" target="_blank">RANK.EQ 函数</a> | 返回某数字在一列数字中的排名 |
| <a href="https://support.office.com/article/RATE-function-9f665657-4a7e-4bb7-a030-83fc59e748ce" target="_blank">RATE 函数</a> | 返回年金的定期利率 |
| <a href="https://support.office.com/article/RECEIVED-function-7a3f8b93-6611-4f81-8576-828312c9b5e5" target="_blank">RECEIVED 函数</a> | 返回完全投资型债券到期收回的金额 |
| <a href="https://support.office.com/article/REPLACE-REPLACEB-functions-8d799074-2425-4a8a-84bc-82472868878a" target="_blank">REPLACE、REPLACEB 函数</a> | 替换文本中的字符 |
| <a href="https://support.office.com/article/REPT-function-04c4d778-e712-43b4-9c15-d656582bb061" target="_blank">REPT 函数</a> | 以给定的次数重复文本 |
| <a href="https://support.office.com/article/RIGHT-RIGHTB-functions-240267ee-9afa-4639-a02b-f19e1786cf2f" target="_blank">RIGHT、RIGHTB 函数</a> | 返回一个文本值的最右端字符 |
| <a href="https://support.office.com/article/ROMAN-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5" target="_blank">ROMAN 函数</a> | 将阿拉伯数字转换为文本形式的罗马数字 |
| <a href="https://support.office.com/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c" target="_blank">ROUND 函数</a> | 将数字舍入到指定位数 |
| <a href="https://support.office.com/article/ROUNDDOWN-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53" target="_blank">ROUNDDOWN 函数</a> | 将数字向零的方向向下舍入 |
| <a href="https://support.office.com/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7" target="_blank">ROUNDUP 函数</a> | 将数字向远离零的方向向上舍入 |
| <a href="https://support.office.com/article/ROWS-function-b592593e-3fc2-47f2-bec1-bda493811597" target="_blank">ROWS 函数</a> | 返回引用中的行数 |
| <a href="https://support.office.com/article/RRI-function-6f5822d8-7ef1-4233-944c-79e8172930f4" target="_blank">RRI 函数</a> | 返回某项投资增长的等效利率 |
| <a href="https://support.office.com/article/SEC-function-ff224717-9c87-4170-9b58-d069ced6d5f7" target="_blank">SEC 函数</a> | 返回一个角度的正割值 |
| <a href="https://support.office.com/article/SECH-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f" target="_blank">SECH 函数</a> | 返回一个角度的双曲正割值 |
| <a href="https://support.office.com/article/SECOND-function-740d1cfc-553c-4099-b668-80eaa24e8af1" target="_blank">SECOND 函数</a> | 将序列号转换为秒 |
| <a href="https://support.office.com/article/SERIESSUM-function-a3ab25b5-1093-4f5b-b084-96c49087f637" target="_blank">SERIESSUM 函数</a> | 返回基于以下公式的幂级数之和 |
| <a href="https://support.office.com/article/SHEET-function-44718b6f-8b87-47a1-a9d6-b701c06cff24" target="_blank">SHEET 函数</a> | 返回引用的工作表的工作表编号 |
| <a href="https://support.office.com/article/SHEETS-function-770515eb-e1e8-45ce-8066-b557e5e4b80b" target="_blank">SHEETS 函数</a> | 返回引用中的工作表数 |
| <a href="https://support.office.com/article/SIGN-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8" target="_blank">SIGN 函数</a> | 返回数值的符号 |
| <a href="https://support.office.com/article/SIN-function-cf0e3432-8b9e-483c-bc55-a76651c95602" target="_blank">SIN 函数</a> | 返回给定角的正弦值 |
| <a href="https://support.office.com/article/SINH-function-1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7" target="_blank">SINH 函数</a> | 返回某一数字的双曲正弦值 |
| <a href="https://support.office.com/article/SKEW-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa" target="_blank">SKEW 函数</a> | 返回一个分布的不对称度 |
| <a href="https://support.office.com/article/SKEWP-function-76530a5c-99b9-48a1-8392-26632d542fcb" target="_blank">SKEW.P 函数</a> | 基于总体返回一个分布的不对称度：用来体现某一分布相对其平均值的不对称程度 |
| <a href="https://support.office.com/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8" target="_blank">SLN 函数</a> | 返回某项资产一个周期的直线折旧值 |
| <a href="https://support.office.com/article/SMALL-function-17da8222-7c82-42b2-961b-14c45384df07" target="_blank">SMALL 函数</a> | 返回数据集中第 k 个最小值 |
| <a href="https://support.office.com/article/SQRT-function-654975c2-05c4-4831-9a24-2c65e4040fdf" target="_blank">SQRT 函数</a> | 返回正平方根 |
| <a href="https://support.office.com/article/SQRTPI-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4" target="_blank">SQRTPI 函数</a> | 返回（数字 * pi）的平方根 |
| <a href="https://support.office.com/article/STANDARDIZE-function-81d66554-2d54-40ec-ba83-6437108ee775" target="_blank">STANDARDIZE 函数</a> | 返回正态分布概率值 |
| <a href="https://support.office.com/article/STDEVP-function-6e917c05-31a0-496f-ade7-4f4e7462f285" target="_blank">STDEV.P 函数</a> | 基于整个样本总体计算标准偏差 |
| <a href="https://support.office.com/article/STDEVS-function-7d69cf97-0c1f-4acf-be27-f3e83904cc23" target="_blank">STDEV.S 函数</a> | 基于样本估计标准偏差 |
| <a href="https://support.office.com/article/STDEVA-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d" target="_blank">STDEVA 函数</a> | 基于样本估计标准偏差，包括数字、文本和逻辑值 |
| <a href="https://support.office.com/article/STDEVPA-function-5578d4d6-455a-4308-9991-d405afe2c28c" target="_blank">STDEVPA 函数</a> | 基于整个样本总体计算标准偏差，包括数字、文本和逻辑值 |
| <a href="https://support.office.com/article/SUBSTITUTE-function-6434944e-a904-4336-a9b0-1e58df3bc332" target="_blank">SUBSTITUTE 函数</a> | 在文本串中用新文本替换旧文本。 |
| <a href="https://support.office.com/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939" target="_blank">SUBTOTAL 函数</a> | 返回一个数据列表或数据库的分类汇总 |
| <a href="https://support.office.com/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89" target="_blank">SUM 函数</a> | 对参数求和 |
| <a href="https://support.office.com/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b" target="_blank">SUMIF 函数</a> | 根据给定的标准，对指定的单元格求和 |
| <a href="https://support.office.com/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b" target="_blank">SUMIFS 函数</a> | 对区域中满足多个条件的单元格求和 |
| <a href="https://support.office.com/article/SUMSQ-function-e3313c02-51cc-4963-aae6-31442d9ec307" target="_blank">SUMSQ 函数</a> | 返回所有参数的平方和 |
| <a href="https://support.office.com/article/SYD-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27" target="_blank">SYD 函数</a> | 返回某项资产在指定期间的年限总额折旧。 |
| <a href="https://support.office.com/article/T-function-fb83aeec-45e7-4924-af95-53e073541228" target="_blank">T 函数</a> | 将其参数转换为文本 |
| <a href="https://support.office.com/article/TDIST-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2" target="_blank">T.DIST 函数</a> | 返回学生 t 分布的百分点（概率） |
| <a href="https://support.office.com/article/TDIST2T-function-198e9340-e360-4230-bd21-f52f22ff5c28" target="_blank">T.DIST.2T 函数</a> | 返回学生 t 分布的百分点（概率） |
| <a href="https://support.office.com/article/TDISTRT-function-20a30020-86f9-4b35-af1f-7ef6ae683eda" target="_blank">T.DIST.RT 函数</a> | 返回学生的 t 分布 |
| <a href="https://support.office.com/article/TINV-function-2908272b-4e61-4942-9df9-a25fec9b0e2e" target="_blank">T.INV 函数</a> | 返回作为概率和自由度函数的学生 t 分布的 t 值 |
| <a href="https://support.office.com/article/TINV2T-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17" target="_blank">T.INV.2T 函数</a> | 返回学生 t 分布的反函数 |
| <a href="https://support.office.com/article/TAN-function-08851a40-179f-4052-b789-d7f699447401" target="_blank">TAN 函数</a> | 返回一个数字的正切值 |
| <a href="https://support.office.com/article/TANH-function-017222f0-a0c3-4f69-9787-b3202295dc6c" target="_blank">TANH 函数</a> | 返回一个数字的双曲正切值 |
| <a href="https://support.office.com/article/TBILLEQ-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c" target="_blank">TBILLEQ 函数</a> | 返回短期国库券的等价债券收益 |
| <a href="https://support.office.com/article/TBILLPRICE-function-eacca992-c29d-425a-9eb8-0513fe6035a2" target="_blank">TBILLPRICE 函数</a> | 返回每张票面为 100 元的短期国库券的现价 |
| <a href="https://support.office.com/article/TBILLYIELD-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba" target="_blank">TBILLYIELD 函数</a> | 返回短期国库券的收益 |
| <a href="https://support.office.com/article/TEXT-function-20d5ac4d-7b94-49fd-bb38-93d29371225c" target="_blank">TEXT 函数</a> | 设置数字格式并将其转换为文本 |
| <a href="https://support.office.com/article/TIME-function-9a5aff99-8f7d-4611-845e-747d0b8d5457" target="_blank">TIME 函数</a> | 返回特定时间的序列号 |
| <a href="https://support.office.com/article/TIMEVALUE-function-0b615c12-33d8-4431-bf3d-f3eb6d186645" target="_blank">TIMEVALUE 函数</a> | 将以文本表达的时间转换为序列号 |
| <a href="https://support.office.com/article/TODAY-function-5eb3078d-a82c-4736-8930-2f51a028fdd9" target="_blank">TODAY 函数</a> | 返回当前日期的序列号 |
| <a href="https://support.office.com/article/TRIM-function-410388fa-c5df-49c6-b16c-9e5630b479f9" target="_blank">TRIM 函数</a> | 从文本中删除空格 |
| <a href="https://support.office.com/article/TRIMMEAN-function-d90c9878-a119-4746-88fa-63d988f511d3" target="_blank">TRIMMEAN 函数</a> | 返回数据集内部的平均值 |
| <a href="https://support.office.com/article/TRUE-function-7652c6e3-8987-48d0-97cd-ef223246b3fb" target="_blank">TRUE 函数</a> | 返回逻辑值 `TRUE` |
| <a href="https://support.office.com/article/TRUNC-function-8b86a64c-3127-43db-ba14-aa5ceb292721" target="_blank">TRUNC 函数</a> | 将数字截断为整数 |
| <a href="https://support.office.com/article/TYPE-function-45b4e688-4bc3-48b3-a105-ffa892995899" target="_blank">TYPE 函数</a> | 返回一个指示数值数据类型的数字 |
| <a href="https://support.office.com/article/UNICHAR-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8" target="_blank">UNICHAR 函数</a> | 返回给定数值引用的 Unicode 字符 |
| <a href="https://support.office.com/article/UNICODE-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f" target="_blank">UNICODE 函数</a> | 返回与文本的第一个字符相对应的数字（码位） |
| <a href="https://support.office.com/article/UPPER-function-c11f29b3-d1a3-4537-8df6-04d0049963d6" target="_blank">UPPER 函数</a> | 将文本转换为大写 |
| <a href="https://support.office.com/article/VALUE-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2" target="_blank">VALUE 函数</a> | 将文本参数转换为数字 |
| <a href="https://support.office.com/article/VARP-function-73d1285c-108c-4843-ba5d-a51f90656f3a" target="_blank">VAR.P 函数</a> | 基于整个样本总体计算方差 |
| <a href="https://support.office.com/article/VARS-function-913633de-136b-449d-813e-65a00b2b990b" target="_blank">VAR.S 函数</a> | 基于样本估计方差 |
| <a href="https://support.office.com/article/VARA-function-3de77469-fa3a-47b4-85fd-81758a1e1d07" target="_blank">VARA 函数</a> | 基于样本估计方差，包括数字、文本和逻辑值 |
| <a href="https://support.office.com/article/VARPA-function-59a62635-4e89-4fad-88ac-ce4dc0513b96" target="_blank">VARPA 函数</a> | 基于整个样本总体计算方差，包括数字、文本和逻辑值 |
| <a href="https://support.office.com/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73" target="_blank">VDB 函数</a> | 使用余额递减法返回指定周期或部分周期内某项资产的折旧值 |
| <a href="https://support.office.com/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1" target="_blank">VLOOKUP 函数</a> | 查找数组的首列并在行间移动以返回单元格的值 |
| <a href="https://support.office.com/article/WEEKDAY-function-60e44483-2ed1-439f-8bd0-e404c190949a" target="_blank">WEEKDAY 函数</a> | 将序列号转换为一周中的某一天 |
| <a href="https://support.office.com/article/WEEKNUM-function-e5c43a03-b4ab-426c-b411-b18c13c75340" target="_blank">WEEKNUM 函数</a> | 将序列号转换为代表一年中第几周的数字 |
| <a href="https://support.office.com/article/WEIBULLDIST-function-4e783c39-9325-49be-bbc9-a83ef82b45db" target="_blank">WEIBULL.DIST 函数</a> | 返回 Weibull 分布 |
| <a href="https://support.office.com/article/WORKDAY-function-f764a5b7-05fc-4494-9486-60d494efbf33" target="_blank">WORKDAY 函数</a> | 返回在指定的若干个工作日之前/之后的日期（一串数字） |
| <a href="https://support.office.com/article/WORKDAYINTL-function-a378391c-9ba7-4678-8a39-39611a9bf81d" target="_blank">WORKDAY.INTL 函数</a> | 返回在指定的若干个工作日之前/之后的日期（一串数字），其中使用参数来指示哪些以及多少天为周末 |
| <a href="https://support.office.com/article/XIRR-function-de1242ec-6477-445b-b11b-a303ad9adc9d" target="_blank">XIRR 函数</a> | 返回一组现金流的内部收益率，这些现金流不一定定期发生 |
| <a href="https://support.office.com/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7" target="_blank">XNPV 函数</a> | 返回一组现金流的净现值，这些现金流不一定定期发生 |
| <a href="https://support.office.com/article/XOR-function-1548d4c2-5e47-4f77-9a92-0533bba14f37" target="_blank">XOR 函数</a> | 返回所有参数的逻辑“异或”值 |
| <a href="https://support.office.com/article/YEAR-function-c64f017a-1354-490d-981f-578e8ec8d3b9" target="_blank">YEAR 函数</a> | 将序列号转换为年 |
| <a href="https://support.office.com/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8" target="_blank">YEARFRAC 函数</a> | 返回表示 start_date 和 end_date 之间的天数占一年总天数的比值 |
| <a href="https://support.office.com/article/YIELD-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe" target="_blank">YIELD 函数</a> | 返回定期支付利息的债券的收益 |
| <a href="https://support.office.com/article/YIELDDISC-function-a9dbdbae-7dae-46de-b995-615faffaaed7" target="_blank">YIELDDISC 函数</a> | 返回已贴现债券的年收益；例如，短期国库券 |
| <a href="https://support.office.com/article/YIELDMAT-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f" target="_blank">YIELDMAT 函数</a> | 返回到期付息的债券的年收益 |
| <a href="https://support.office.com/article/ZTEST-function-d633d5a3-2031-4614-a016-92180ad82bee" target="_blank">Z.TEST 函数</a> | 返回 z 检验的收尾概率值 |

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)
- [函数类（适用于 Excel 的 JavaScript API）](/javascript/api/excel/excel.functions)
- [工作簿函数对象（适用于 Excel 的 JavaScript API）](/javascript/api/excel/excel.workbook#functions)
