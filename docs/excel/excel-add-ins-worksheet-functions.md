---
title: 使用 Excel JavaScript API 调用内置 Excel 工作表函数
description: 了解如何在 JavaScript API 等`VLOOKUP`Excel调用`SUM`内置函数，Excel JavaScript API。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: cc7622b642720a8cb8f80ad553600fd22ac7c25c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744072"
---
# <a name="call-built-in-excel-worksheet-functions"></a>调用内置 Excel 工作表函数

本文介绍了如何使用 Excel JavaScript API 调用内置 Excel 工作表函数（如 `VLOOKUP` 和 `SUM`）。 其中还收录了可以使用 Excel JavaScript API 调用的内置 Excel 工作表函数的完整列表。

> [!NOTE]
> 若要了解如何使用 Excel JavaScript API 在 Excel 中创建 *自定义函数*，请参阅 [在 Excel 中创建自定义函数](custom-functions-overview.md)。

## <a name="calling-a-worksheet-function"></a>创建工作表函数

下面的代码片段展示了如何调用工作表函数，其中 `sampleFunction()` 是占位符，应将它替换为要调用的函数名称和函数需要使用的输入参数。 工作表 `value` 函数 `FunctionResult` 返回的对象的属性包含指定函数的结果。 如以下示例所示，您`load``value`必须先使用 对象的 属性`FunctionResult`，然后才能读取该对象。 在此示例中，函数结果被直接写入控制台。

```js
await Excel.run(async (context) => {
    let functionResult = context.workbook.functions.sampleFunction();
    functionResult.load('value');

    await context.sync();
    console.log('Result of the function: ' + functionResult.value);
});
```

> [!TIP]
> 有关可以使用 Excel JavaScript API 调用的函数列表，请参阅本文的[支持的工作表函数](#supported-worksheet-functions)部分。

## <a name="sample-data"></a>示例数据

下图展示了 Excel 工作表中的表格，其中包含三个月内各种工具的销售数据。 表格中的每个数字均表示具体工具在特定月份中的销售件数。 接下来的两个示例展示了如何向此类数据应用内置工作表函数。

![在 11 月、12 月和 1 月Excel一个月份中，一个"扳手"和"扳手"的销售数据的屏幕截图。](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>示例 1：单函数

下面的代码示例向前述示例数据应用 `VLOOKUP` 函数，以确定 11 月售出的扳手数。

```js
await Excel.run(async (context) => {
    let range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    let unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    await context.sync();
    console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
});
```

## <a name="example-2-nested-functions"></a>示例 2：嵌套函数

下面的代码示例向前述示例数据应用 `VLOOKUP` 函数，以分别确定 11 月和 12 月售出的扳手数。然后，应用 `SUM` 函数，以计算这两个月售出的扳手总数。

如此示例所示，如果一个或多个函数调用嵌套在另一个函数调用中，只需对随后要读取的最终结果（在此示例中为 `load`）执行 `sumOfTwoLookups` 操作即可。 系统会计算所有中间结果（在此示例中，为每个 `VLOOKUP` 函数的结果），并根据这些中间结果计算最终结果。

```js
await Excel.run(async (context) => {
    let range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    let sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false),
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    await context.sync();
    console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
});
```

## <a name="supported-worksheet-functions"></a>支持的工作表函数

可以使用 Excel JavaScript API 调用以下内置 Excel 工作表函数。

| 功能 | 说明 |
|:---------------|:-----------|
| [ABS 函数](https://support.microsoft.com/office/3420200f-5628-4e8c-99da-c99d7c87713c) | 返回数字的绝对值 |
| [ACCRINT 函数](https://support.microsoft.com/office/fe45d089-6722-4fb3-9379-e1f911d8dc74) | 返回定期支付利息的债券的应计利息 |
| [ACCRINTM 函数](https://support.microsoft.com/office/f62f01f9-5754-4cc4-805b-0e70199328a7) | 返回在到期日支付利息的债券的应计利息 |
| [ACOS 函数](https://support.microsoft.com/office/cb73173f-d089-4582-afa1-76e5524b5d5b) | 返回一个数的反余弦值 |
| [ACOSH 函数](https://support.microsoft.com/office/e3992cc1-103f-4e72-9f04-624b9ef5ebfe) | 返回一个数的反双曲余弦值 |
| [ACOT 函数](https://support.microsoft.com/office/dc7e5008-fe6b-402e-bdd6-2eea8383d905) | 返回一个数的反余切值 |
| [ACOTH 函数](https://support.microsoft.com/office/cc49480f-f684-4171-9fc5-73e4e852300f) | 返回一个数的双曲反余切值 |
| [AMORDEGRC 函数](https://support.microsoft.com/office/a14d0ca1-64a4-42eb-9b3d-b0dededf9e51) | 通过使用折旧系数，返回每个会计期间的折旧值 |
| [AMORLINC 函数](https://support.microsoft.com/office/7d417b45-f7f5-4dba-a0a5-3451a81079a8) | 返回每个会计期间的折旧值 |
| [AND 函数](https://support.microsoft.com/office/5f19b2e8-e1df-4408-897a-ce285a19e9d9) | 如果所有参数都为 true，返回 `TRUE` |
| [ARABIC 函数](https://support.microsoft.com/office/9a8da418-c17b-4ef9-a657-9370a30a674f) | 将罗马数字转换为阿拉伯数字 |
| [AREAS 函数](https://support.microsoft.com/office/8392ba32-7a41-43b3-96b0-3695d2ec6152) | 返回引用中包含的区域个数 |
| [ASC 函数](https://support.microsoft.com/office/0b6abf1c-c663-4004-a964-ebc00b723266) | 将字符串中的全角（双字节）英文字母或片假名更改为半角（单字节）字符 |
| [ASIN 函数](https://support.microsoft.com/office/81fb95e5-6d6f-48c4-bc45-58f955c6d347) | 返回一个数的反正弦值 |
| [ASINH 函数](https://support.microsoft.com/office/4e00475a-067a-43cf-926a-765b0249717c) | 返回一个数的反双曲正弦值 |
| [ATAN 函数](https://support.microsoft.com/office/50746fa8-630a-406b-81d0-4a2aed395543) | 返回一个数的反正切值 |
| [ATAN2 函数](https://support.microsoft.com/office/c04592ab-b9e3-4908-b428-c96b3a565033) | 返回从 x 坐标和 y 坐标的反正切值 |
| [ATANH 函数](https://support.microsoft.com/office/3cd65768-0de7-4f1d-b312-d01c8c930d90) | 返回某一数字的反双曲正切值 |
| [AVEDEV 函数](https://support.microsoft.com/office/58fe8d65-2a84-4dc7-8052-f3f87b5c6639) | 返回一组数据点到其算术平均值的绝对偏差的平均值 |
| [AVERAGE 函数](https://support.microsoft.com/office/047bac88-d466-426c-a32b-8f33eb960cf6) | 返回其参数的平均值 |
| [AVERAGEA 函数](https://support.microsoft.com/office/f5f84098-d453-4f4c-bbba-3d2c66356091) | 返回其参数的平均值，包括数字、文本和逻辑值 |
| [AVERAGEIF 函数](https://support.microsoft.com/office/faec8e2e-0dec-4308-af69-f5576d8ac642) | 返回区域内满足给定条件的所有单元格的平均值（算术平均值） |
| [AVERAGEIFS 函数](https://support.microsoft.com/office/48910c45-1fc0-4389-a028-f7c5c3001690) | 返回满足多个条件的所有单元格的平均值（算术平均值） |
| [BAHTTEXT 函数](https://support.microsoft.com/office/5ba4d0b4-abd3-4325-8d22-7a92d59aab9c) | 使用 ß（铢）货币格式将数字转换为文本 |
| [BASE 函数](https://support.microsoft.com/office/2ef61411-aee9-4f29-a811-1c42456c6342) | 将数字转换成具有给定基数的文本表示形式 |
| [BESSELI 函数](https://support.microsoft.com/office/8d33855c-9a8d-444b-98e0-852267b1c0df) | 返回修正的贝塞耳函数 In(x) |
| [BESSELJ 函数](https://support.microsoft.com/office/839cb181-48de-408b-9d80-bd02982d94f7) | 返回贝塞耳函数 Jn(x) |
| [BESSELK 函数](https://support.microsoft.com/office/606d11bc-06d3-4d53-9ecb-2803e2b90b70) | 返回修正的贝塞耳函数 Kn(x) |
| [BESSELY 函数](https://support.microsoft.com/office/f3a356b3-da89-42c3-8974-2da54d6353a2) | 返回贝赛耳函数 Yn(x) |
| [BETA.DIST 函数](https://support.microsoft.com/office/11188c9c-780a-42c7-ba43-9ecb5a878d31) | 返回 beta 累积分布函数 |
| [BETA.INV 函数](https://support.microsoft.com/office/e84cb8aa-8df0-4cf6-9892-83a341d252eb) | 返回指定的 beta 分布累积分布函数的反函数 |
| [BIN2DEC 函数](https://support.microsoft.com/office/63905b57-b3a0-453d-99f4-647bb519cd6c) | 将二进制数转换为十进制 |
| [BIN2HEX 函数](https://support.microsoft.com/office/0375e507-f5e5-4077-9af8-28d84f9f41cc) | 将二进制数转换为十六进制 |
| [BIN2OCT 函数](https://support.microsoft.com/office/0a4e01ba-ac8d-4158-9b29-16c25c4c23fd) | 将二进制数转换为八进制 |
| [BINOM.DIST 函数](https://support.microsoft.com/office/c5ae37b6-f39c-4be2-94c2-509a1480770c) | 返回一元二项式分布的概率 |
| [BINOM.DIST.RANGE 函数](https://support.microsoft.com/office/17331329-74c7-4053-bb4c-6653a7421595) | 返回使用二项式分布的试验结果的概率 |
| [BINOM.INV 函数](https://support.microsoft.com/office/80a0370c-ada6-49b4-83e7-05a91ba77ac9) | 返回一个数值，它是使得累积二项式分布的函数值小于或等于临界值的最小整数 |
| [BITAND 函数](https://support.microsoft.com/office/8a2be3d7-91c3-4b48-9517-64548008563a) | 返回两个数字的“按位与” |
| [BITLSHIFT 函数](https://support.microsoft.com/office/c55bb27e-cacd-4c7c-b258-d80861a03c9c) | 返回按照 shift_amount 位数左移后得到的数值 |
| [BITOR 函数](https://support.microsoft.com/office/f6ead5c8-5b98-4c9e-9053-8ad5234919b2) | 返回 2 个数字的按位“或” |
| [BITRSHIFT 函数](https://support.microsoft.com/office/274d6996-f42c-4743-abdb-4ff95351222c) | 返回按照 shift_amount 位数右移后得到的数值 |
| [BITXOR Function](https://support.microsoft.com/office/c81306a1-03f9-4e89-85ac-b86c3cba10e4) | 返回两个数字的按位“异或”值 |
| [CEILING。MATH、ECMA_CEILING 函数](https://support.microsoft.com/office/80f95d2f-b499-4eee-9f16-f795a8e306c8) | 将数值向上舍入为最接近的整数或最接近的基数的倍数 |
| [CEILING.PRECISE 函数](https://support.microsoft.com/office/f366a774-527a-4c92-ba49-af0a196e66cb) | 将数值四舍五入到最接近的整数或最接近的基数的倍数。不论数字是否带有符号，都将数字向上舍入。 |
| [CHAR 函数](https://support.microsoft.com/office/bbd249c8-b36e-4a91-8017-1c133f9b837a) | 返回由代码数字指定的字符 |
| [CHISQ.DIST 函数](https://support.microsoft.com/office/8486b05e-5c05-4942-a9ea-f6b341518732) | 返回累积 beta 分布的概率密度函数 |
| [CHISQ.DIST.RT 函数](https://support.microsoft.com/office/dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2) | 返回 χ2 分布的收尾概率 |
| [CHISQ.INV 函数](https://support.microsoft.com/office/400db556-62b3-472d-80b3-254723e7092f) | 返回累积 beta 分布的概率密度函数 |
| [CHISQ.INV.RT 函数](https://support.microsoft.com/office/435b5ed8-98d5-4da6-823f-293e2cbc94fe) | 返回 χ2 分布的收尾概率的反函数 |
| [CHOOSE 函数](https://support.microsoft.com/office/fc5c184f-cb62-4ec7-a46e-38653b98f5bc) | 从值列表中选择一个值 |
| [CLEAN 函数](https://support.microsoft.com/office/26f3d7c5-475f-4a9c-90e5-4b8ba987ba41) | 删除文本中的所有非打印字符 |
| [CODE 函数](https://support.microsoft.com/office/c32b692b-2ed0-4a04-bdd9-75640144b928) | 返回文本字符串中第一个字符的数字代码 |
| [COLUMNS 函数](https://support.microsoft.com/office/4e8e7b4e-e603-43e8-b177-956088fa48ca) | 返回引用中的列数 |
| [COMBIN 函数](https://support.microsoft.com/office/12a3f276-0a21-423a-8de6-06990aaf638a) | 返回给定数目对象的组合数 |
| [COMBINA 函数](https://support.microsoft.com/office/efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d) | 返回给定项数的组合数（包含重复项） |
| [COMPLEX 函数](https://support.microsoft.com/office/f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128) | 将实部系数和虚部系数转换为复数 |
| [CONCATENATE 函数](https://support.microsoft.com/office/8f8ae884-2ca8-4f7a-b093-75d702bea31d) | 将几个文本项合并为一个文本项 |
| [CONFIDENCE.NORM 函数](https://support.microsoft.com/office/7cec58a6-85bb-488d-91c3-63828d4fbfd4) | 返回总体平均数的置信区间 |
| [CONFIDENCE.T 函数](https://support.microsoft.com/office/e8eca395-6c3a-4ba9-9003-79ccc61d3c53) | 使用学生 t 分布返回总体平均数的置信区间 |
| [CONVERT 函数](https://support.microsoft.com/office/d785bef1-808e-4aac-bdcd-666c810f9af2) | 将数字从一种度量体系转换为另一种度量体系 |
| [COS 函数](https://support.microsoft.com/office/0fb808a5-95d6-4553-8148-22aebdce5f05) | 返回一个数的余弦值 |
| [COSH 函数](https://support.microsoft.com/office/e460d426-c471-43e8-9540-a57ff3b70555) | 返回一个数字的双曲余弦值 |
| [COT 函数](https://support.microsoft.com/office/c446f34d-6fe4-40dc-84f8-cf59e5f5e31a) | 返回一个角度的余切值 |
| [COTH 函数](https://support.microsoft.com/office/2e0b4cb6-0ba0-403e-aed4-deaa71b49df5) | 返回一个数字的双曲余切值 |
| [COUNT 函数](https://support.microsoft.com/office/a59cd7fc-b623-4d93-87a4-d23bf411294c) | 计算参数表中的数字个数 |
| [COUNTA 函数](https://support.microsoft.com/office/7dc98875-d5c1-46f1-9a82-53f3219e2509) | 计算参数列表中值的数量 |
| [COUNTBLANK 函数](https://support.microsoft.com/office/6a92d772-675c-4bee-b346-24af6bd3ac22) | 计算在一定范围内的空单元格数量 |
| [COUNTIF 函数](https://support.microsoft.com/office/e0de10c6-f885-4e71-abb4-1f464816df34) | 计算某个区域中满足给定条件的单元格数目 |
| [COUNTIFS 函数](https://support.microsoft.com/office/dda3dc6e-f74e-4aee-88bc-aa8c2a866842) | 计算某个区域中满足多个条件的单元格数目 |
| [COUPDAYBS 函数](https://support.microsoft.com/office/eb9a8dfb-2fb2-4c61-8e5d-690b320cf872) | 返回从票息期开始到结算日之间的天数 |
| [COUPDAYS 函数](https://support.microsoft.com/office/cc64380b-315b-4e7b-950c-b30b0a76f671) | 返回包含结算日的票息期的天数 |
| [COUPDAYSNC 函数](https://support.microsoft.com/office/5ab3f0b2-029f-4a8b-bb65-47d525eea547) | 返回从结算日到下一票息支付日之间的天数 |
| [COUPNCD 函数](https://support.microsoft.com/office/fd962fef-506b-4d9d-8590-16df5393691f) | 返回结算日后的下一票息支付日 |
| [COUPNUM 函数](https://support.microsoft.com/office/a90af57b-de53-4969-9c99-dd6139db2522) | 返回结算日与到期日之间可支付的票息数 |
| [COUPPCD 函数](https://support.microsoft.com/office/2eb50473-6ee9-4052-a206-77a9a385d5b3) | 返回结算日前的上一票息支付日 |
| [CSC 函数](https://support.microsoft.com/office/07379361-219a-4398-8675-07ddc4f135c1) | 返回一个角度的余割值 |
| [CSCH 函数](https://support.microsoft.com/office/f58f2c22-eb75-4dd6-84f4-a503527f8eeb) | 返回一个角度的双曲余割值 |
| [CUMIPMT 函数](https://support.microsoft.com/office/61067bb0-9016-427d-b95b-1a752af0e606) | 返回两个付款期之间为贷款累积支付的利息 |
| [CUMPRINC 函数](https://support.microsoft.com/office/94a4516d-bd65-41a1-bc16-053a6af4c04d) | 返回两个付款期之间为贷款累积支付的本金 |
| [DATE 函数](https://support.microsoft.com/office/e36c0c8c-4104-49da-ab83-82328b832349) | 返回特定日期的序列号 |
| [DATEVALUE 函数](https://support.microsoft.com/office/df8b07d4-7761-4a93-bc33-b7471bbff252) | 将以文本表达的日期转换为序列号 |
| [DAVERAGE 函数](https://support.microsoft.com/office/a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee) | 返回所选数据库条目的平均值 |
| [DAY 函数](https://support.microsoft.com/office/8a7d1cbb-6c7d-4ba1-8aea-25c134d03101) | 将序列号转换为月份中的某一天 |
| [DAYS 函数](https://support.microsoft.com/office/57740535-d549-4395-8728-0f07bff0b9df) | 返回两个日期间相差的天数 |
| [DAYS360 函数](https://support.microsoft.com/office/b9a509fd-49ef-407e-94df-0cbda5718c2a) | 按每年 360 天计算两个日期间相差的天数 |
| [DB 函数](https://support.microsoft.com/office/354e7d28-5f93-4ff1-8a52-eb4ee549d9d7) | 使用固定余额递减法返回指定周期内某项资产的折旧值 |
| [DBCS 函数](https://support.microsoft.com/office/a4025e73-63d2-4958-9423-21a24794c9e5) | 将字符串中的半角（单字节）英文字母或片假名更改为全角（双字节）字符 |
| [DCOUNT 函数](https://support.microsoft.com/office/c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1) | 计算数据库中包含数字的单元格数量 |
| [DCOUNTA 函数](https://support.microsoft.com/office/00232a6d-5a66-4a01-a25b-c1653fda1244) | 计算数据库中的非空单元格的数量 |
| [DDB 函数](https://support.microsoft.com/office/519a7a37-8772-4c96-85c0-ed2c209717a5) | 使用双倍余额递减法或其他指定方法返回某项资产在指定周期内的折旧值 |
| [DEC2BIN 函数](https://support.microsoft.com/office/0f63dd0e-5d1a-42d8-b511-5bf5c6d43838) | 将十进制数转换为二进制 |
| [DEC2HEX 函数](https://support.microsoft.com/office/6344ee8b-b6b5-4c6a-a672-f64666704619) | 将十进制数转换为十六进制 |
| [DEC2OCT 函数](https://support.microsoft.com/office/c9d835ca-20b7-40c4-8a9e-d3be351ce00f) | 将十进制数转换为八进制 |
| [DECIMAL 函数](https://support.microsoft.com/office/ee554665-6176-46ef-82de-0a283658da2e) | 按给定基数将数字的文本表示形式转换成十进制数 |
| [DEGREES 函数](https://support.microsoft.com/office/4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1) | 将弧度转换为角度 |
| [DELTA 函数](https://support.microsoft.com/office/2f763672-c959-4e07-ac33-fe03220ba432) | 测试两个值是否相等 |
| [DEVSQ 函数](https://support.microsoft.com/office/8b739616-8376-4df5-8bd0-cfe0a6caf444) | 返回偏差平方和 |
| [DGET 函数](https://support.microsoft.com/office/455568bf-4eef-45f7-90f0-ec250d00892e) | 从数据库中提取符合指定条件的单个记录 |
| [DISC 函数](https://support.microsoft.com/office/71fce9f3-3f05-4acf-a5a3-eac6ef4daa53) | 返回债券的贴现率 |
| [DMAX 函数](https://support.microsoft.com/office/f4e8209d-8958-4c3d-a1ee-6351665d41c2) | 返回所选数据库条目中的最大值 |
| [DMIN 函数](https://support.microsoft.com/office/4ae6f1d9-1f26-40f1-a783-6dc3680192a3) | 返回所选数据库条目中的最小值 |
| [DOLLAR、USDOLLAR 函数](https://support.microsoft.com/office/a6cd05d9-9740-4ad3-a469-8109d18ff611) | 使用 $（美元）货币格式将数字转换为文本 |
| [DOLLARDE 函数](https://support.microsoft.com/office/db85aab0-1677-428a-9dfd-a38476693427) | 将以分数表示的货币值转换为以小数表示的货币值 |
| [DOLLARFR 函数](https://support.microsoft.com/office/0835d163-3023-4a33-9824-3042c5d4f495) | 将以小数表示的货币值转换为以分数表示的货币值 |
| [DPRODUCT 函数](https://support.microsoft.com/office/4f96b13e-d49c-47a7-b769-22f6d017cb31) | 将与数据库中的条件匹配的记录的特定字段中的值相乘 |
| [DSTDEV 函数](https://support.microsoft.com/office/026b8c73-616d-4b5e-b072-241871c4ab96) | 根据所选数据库条目中的样本估算数据的标准偏差 |
| [DSTDEVP 函数](https://support.microsoft.com/office/04b78995-da03-4813-bbd9-d74fd0f5d94b) | 以数据库选定项作为样本总体，计算数据的标准偏差 |
| [DSUM 函数](https://support.microsoft.com/office/53181285-0c4b-4f5a-aaa3-529a322be41b) | 将数据库中与条件匹配的记录字段列中的数字进行求和 |
| [DURATION 函数](https://support.microsoft.com/office/b254ea57-eadc-4602-a86a-c8e369334038) | 返回定期支付利息的债券的年持续时间 |
| [Dlet 函数](https://support.microsoft.com/office/d6747ca9-99c7-48bb-996e-9d7af00f3ed1) | 根据所选数据库条目中的样本估算数据的方差 |
| [DVARP 函数](https://support.microsoft.com/office/eb0ba387-9cb7-45c8-81e9-0394912502fc) | 以数据库选定项作为样本总体，计算数据的总体方差 |
| [EDATE 函数](https://support.microsoft.com/office/3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) | 返回一串日期，指示起始日期之前/之后的月数 |
| [EFFECT 函数](https://support.microsoft.com/office/910d4e4c-79e2-4009-95e6-507e04f11bc4) | 返回年有效利率 |
| [EOMONTH 函数](https://support.microsoft.com/office/7314ffa1-2bc9-4005-9d66-f49db127d628) | 返回一串日期，表示指定月数之前或之后的月份的最后一天 |
| [ERF 函数](https://support.microsoft.com/office/c53c7e7b-5482-4b6c-883e-56df3c9af349) | 返回误差函数 |
| [ERF.PRECISE 函数](https://support.microsoft.com/office/9a349593-705c-4278-9a98-e4122831a8e0) | 返回误差函数 |
| [ERFC 函数](https://support.microsoft.com/office/736e0318-70ba-4e8b-8d08-461fe68b71b3) | 返回补余误差函数 |
| [ERFC.PRECISE 函数](https://support.microsoft.com/office/e90e6bab-f45e-45df-b2ac-cd2eb4d4a273) | 返回在 x 和无穷大之间集成的补余 ERF 函数 |
| [ERROR.TYPE 函数](https://support.microsoft.com/office/10958677-7c8d-44f7-ae77-b9a9ee6eefaa) | 返回对应于一种错误类型的数字 |
| [EVEN 函数](https://support.microsoft.com/office/197b5f06-c795-4c1e-8696-3c3b8a646cf9) | 将数字向上舍入到最近的偶数 |
| [EXACT 函数](https://support.microsoft.com/office/d3087698-fc15-4a15-9631-12575cf29926) | 检查两个文本值是否相同 |
| [EXP 函数](https://support.microsoft.com/office/c578f034-2c45-4c37-bc8c-329660a63abe) | 返回 e 的 n 次方 |
| [EXPON.DIST 函数](https://support.microsoft.com/office/4c12ae24-e563-4155-bf3e-8b78b6ae140e) | 返回指数分布 |
| [F.DIST 函数](https://support.microsoft.com/office/a887efdc-7c8e-46cb-a74a-f884cd29b25d) | 返回 F 概率分布 |
| [F.DIST.RT 函数](https://support.microsoft.com/office/d74cbb00-6017-4ac9-b7d7-6049badc0520) | 返回 F 概率分布 |
| [F.INV 函数](https://support.microsoft.com/office/0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe) | 返回 F 概率分布的逆函数值 |
| [F.INV.RT 函数](https://support.microsoft.com/office/d371aa8f-b0b1-40ef-9cc2-496f0693ac00) | 返回 F 概率分布的逆函数值 |
| [FACT 函数](https://support.microsoft.com/office/ca8588c2-15f2-41c0-8e8c-c11bd471a4f3) | 返回某数的阶乘 |
| [FACTDOUBLE 函数](https://support.microsoft.com/office/e67697ac-d214-48eb-b7b7-cce2589ecac8) | 返回数字的双阶乘 |
| [FALSE 函数](https://support.microsoft.com/office/2d58dfa5-9c03-4259-bf8f-f0ae14346904) | 返回逻辑值 `FALSE` |
| [FIND、FINDB 函数](https://support.microsoft.com/office/c7912941-af2a-4bdf-a553-d0d89b0a0628) | 在一个文本值中查找另一个（区分大小写） |
| [FISHER 函数](https://support.microsoft.com/office/d656523c-5076-4f95-b87b-7741bf236c69) | 返回 Fisher 变换值 |
| [FISHERINV 函数](https://support.microsoft.com/office/62504b39-415a-4284-a285-19c8e82f86bb) | 返回 Fisher 逆变换值 |
| [FIXED 函数](https://support.microsoft.com/office/ffd5723c-324c-45e9-8b96-e41be2a8274a) | 将数字格式化为具有固定数量的小数的文本 |
| [FLOOR.MATH 函数](https://support.microsoft.com/office/c302b599-fbdb-4177-ba19-2c2b1249a2f5) | 将数字向下舍入为最接近的整数或最接近的基数的倍数 |
| [FLOOR.PRECISE 函数](https://support.microsoft.com/office/f769b468-1452-4617-8dc3-02f842a0702e) | 将数字向下舍入为最接近的整数或最接近的基数的倍数。不论数字是否带有符号，都将数字向下舍入。 |
| [FV 函数](https://support.microsoft.com/office/2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3) | 返回一项投资的未来值 |
| [FVSCHEDULE 函数](https://support.microsoft.com/office/bec29522-bd87-4082-bab9-a241f3fb251d) | 返回在应用一系列复利后，初始本金的终值 |
| [GAMMA 函数](https://support.microsoft.com/office/ce1702b1-cf55-471d-8307-f83be0fc5297) | 返回 Gamma 函数值 |
| [GAMMA.DIST 函数](https://support.microsoft.com/office/9b6f1538-d11c-4d5f-8966-21f6a2201def) | 返回 γ 分布 |
| [GAMMA.INV 函数](https://support.microsoft.com/office/74991443-c2b0-4be5-aaab-1aa4d71fbb18) | 返回 γ 累积分布的反函数 |
| [GAMMALN 函数](https://support.microsoft.com/office/b838c48b-c65f-484f-9e1d-141c55470eb9) | 返回 γ 函数的自然对数 Γ(x) |
| [GAMMALN.PRECISE 函数](https://support.microsoft.com/office/5cdfe601-4e1e-4189-9d74-241ef1caa599) | 返回 γ 函数的自然对数 Γ(x) |
| [GAUSS 函数](https://support.microsoft.com/office/069f1b4e-7dee-4d6a-a71f-4b69044a6b33) | 返回比标准正态累积分布小 0.5 的值 |
| [GCD 函数](https://support.microsoft.com/office/d5107a51-69e3-461f-8e4c-ddfc21b5073a) | 返回最大公约数 |
| [GEOMEAN 函数](https://support.microsoft.com/office/db1ac48d-25a5-40a0-ab83-0b38980e40d5) | 返回几何平均数 |
| [GESTEP 函数](https://support.microsoft.com/office/f37e7d2a-41da-4129-be95-640883fca9df) | 测试某个数字是否大于阈值 |
| [HARMEAN 函数](https://support.microsoft.com/office/5efd9184-fab5-42f9-b1d3-57883a1d3bc6) | 返回调和平均值 |
| [HEX2BIN 函数](https://support.microsoft.com/office/a13aafaa-5737-4920-8424-643e581828c1) | 将十六进制数转换为二进制 |
| [HEX2DEC 函数](https://support.microsoft.com/office/8c8c3155-9f37-45a5-a3ee-ee5379ef106e) | 将十六进制数转换为十进制 |
| [HEX2OCT 函数](https://support.microsoft.com/office/54d52808-5d19-4bd0-8a63-1096a5d11912) | 将十六进制数转换为八进制 |
| [HLOOKUP 函数](https://support.microsoft.com/office/a3034eec-b719-4ba3-bb65-e1ad662ed95f) | 在数组的顶行中查找并返回指定单元格的值 |
| [HOUR 函数](https://support.microsoft.com/office/a3afa879-86cb-4339-b1b5-2dd2d7310ac7) | 将序列号转换为小时 |
| [HYPERLINK 函数](https://support.microsoft.com/office/333c7ce6-c5ae-4164-9c47-7de9b76f577f) | 创建一个快捷方式或链接，以便打开一个存储在网络服务器、内部网或 Internet 上的文档 |
| [HYPGEOM.DIST 函数](https://support.microsoft.com/office/6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf) | 返回超几何分布 |
| [IF 函数](https://support.microsoft.com/office/69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) | 指定要执行的逻辑测试 |
| [IMABS 函数](https://support.microsoft.com/office/b31e73c6-d90c-4062-90bc-8eb351d765a1) | 返回复数的绝对值（模数） |
| [IMAGINARY 函数](https://support.microsoft.com/office/dd5952fd-473d-44d9-95a1-9a17b23e428a) | 返回复数的虚部系数 |
| [IMARGUMENT 函数](https://support.microsoft.com/office/eed37ec1-23b3-4f59-b9f3-d340358a034a) | 返回以弧度表示的角 - 参数 θ |
| [IMCONJUGATE 函数](https://support.microsoft.com/office/2e2fc1ea-f32b-4f9b-9de6-233853bafd42) | 返回复数的共轭复数 |
| [IMCOS 函数](https://support.microsoft.com/office/dad75277-f592-4a6b-ad6c-be93a808a53c) | 返回复数的余弦值 |
| [IMCOSH 函数](https://support.microsoft.com/office/053e4ddb-4122-458b-be9a-457c405e90ff) | 返回复数的双曲余弦值 |
| [IMCOT 函数](https://support.microsoft.com/office/dc6a3607-d26a-4d06-8b41-8931da36442c) | 返回复数的余切值 |
| [IMCSC 函数](https://support.microsoft.com/office/9e158d8f-2ddf-46cd-9b1d-98e29904a323) | 返回复数的余割值 |
| [IMCSCH 函数](https://support.microsoft.com/office/c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9) | 返回复数的双曲余割值 |
| [IMDIV 函数](https://support.microsoft.com/office/a505aff7-af8a-4451-8142-77ec3d74d83f) | 返回两个复数之商 |
| [IMEXP 函数](https://support.microsoft.com/office/c6f8da1f-e024-4c0c-b802-a60e7147a95f) | 返回复数的指数值 |
| [IMLN 函数](https://support.microsoft.com/office/32b98bcf-8b81-437c-a636-6fb3aad509d8) | 返回复数的自然对数 |
| [IMLOG10 函数](https://support.microsoft.com/office/58200fca-e2a2-4271-8a98-ccd4360213a5) | 返回以 10 为底的复数的对数 |
| [IMLOG2 函数](https://support.microsoft.com/office/152e13b4-bc79-486c-a243-e6a676878c51) | 返回以 2 为底的复数的对数 |
| [IMPOWER 函数](https://support.microsoft.com/office/210fd2f5-f8ff-4c6a-9d60-30e34fbdef39) | 返回复数的整数幂 |
| [IMPRODUCT 函数](https://support.microsoft.com/office/2fb8651a-a4f2-444f-975e-8ba7aab3a5ba) | 返回从 2 到 255 个复数的乘积 |
| [IMREAL 函数](https://support.microsoft.com/office/d12bc4c0-25d0-4bb3-a25f-ece1938bf366) | 返回复数的实部系数 |
| [IMSEC 函数](https://support.microsoft.com/office/6df11132-4411-4df4-a3dc-1f17372459e0)IMSEC 函数 | 返回复数的正割值 |
| [IMSECH 函数](https://support.microsoft.com/office/f250304f-788b-4505-954e-eb01fa50903b) | 返回复数的双曲正割值 |
| [IMSIN 函数](https://support.microsoft.com/office/1ab02a39-a721-48de-82ef-f52bf37859f6) | 返回复数的正弦值 |
| [IMSINH 函数](https://support.microsoft.com/office/dfb9ec9e-8783-4985-8c42-b028e9e8da3d) | 返回复数的双曲正弦值 |
| [IMSQRT 函数](https://support.microsoft.com/office/e1753f80-ba11-4664-a10e-e17368396b70) | 返回复数的平方根 |
| [IMSUB 函数](https://support.microsoft.com/office/2e404b4d-4935-4e85-9f52-cb08b9a45054) | 返回两个复数的差值 |
| [IMSUM 函数](https://support.microsoft.com/office/81542999-5f1c-4da6-9ffe-f1d7aaa9457f) | 返回复数的和 |
| [IMTAN 函数](https://support.microsoft.com/office/8478f45d-610a-43cf-8544-9fc0b553a132) | 返回复数的正切值 |
| [INT 函数](https://support.microsoft.com/office/a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) | 将数值向下舍入到最接近的整数 |
| [INTRATE 函数](https://support.microsoft.com/office/5cb34dde-a221-4cb6-b3eb-0b9e55e1316f) | 返回完全投资型债券的利率 |
| [IPmt 函数](https://support.microsoft.com/office/5cce0ad6-8402-4a41-8d29-61a0b054cb6f) | 返回给定期间内投资所支付的利息 |
| [IRR 函数](https://support.microsoft.com/office/64925eaa-9988-495b-b290-3ad0c163c1bc) | 返回一系列现金流的内部收益率 |
| [ISERR 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值是除 #N/A 之外的错误值，返回 `TRUE` |
| [ISERROR 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值是任何错误值，返回 `TRUE` |
| [ISEVEN 函数](https://support.microsoft.com/office/aa15929a-d77b-4fbb-92f4-2f479af55356) | 如果值是偶数，返回 `TRUE` |
| [ISFORMULA 函数](https://support.microsoft.com/office/e4d1355f-7121-4ef2-801e-3839bfd6b1e5) | 如果存在对包含公式的单元格的引用，返回 `TRUE` |
| [ISLOGICAL 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值是逻辑值，返回 `TRUE` |
| [ISNA 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值是 #N/A 错误值，返回 `TRUE` |
| [ISNONTEXT 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值不是文本，返回 `TRUE` |
| [ISNUMBER 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值是数字，返回 `TRUE` |
| [ISO.CEILING 函数](https://support.microsoft.com/office/e587bb73-6cc2-4113-b664-ff5b09859a83) | 将数字向上舍入到最接近的整数或最接近的基数的倍数 |
| [ISODD 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值是奇数，返回 `TRUE` |
| [ISOWEEKNUM 函数](https://support.microsoft.com/office/1c2d0afe-d25b-4ab1-8894-8d0520e90e0e) | 返回一年中给定日期的 ISO 周数的数目 |
| [ISPMT 函数](https://support.microsoft.com/office/fa58adb6-9d39-4ce0-8f43-75399cea56cc) | 计算指定的投资期间支付的利息 |
| [ISREF 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值是引用，返回 `TRUE` |
| [ISTEXT 函数](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | 如果值是文本，返回 `TRUE` |
| [KURT 函数](https://support.microsoft.com/office/bc3a265c-5da4-4dcb-b7fd-c237789095ab) | 返回一组数据的峰值 |
| [LARGE 函数](https://support.microsoft.com/office/3af0af19-1190-42bb-bb8b-01672ec00a64) | 返回数据集中第 k 个最大值 |
| [LCM 函数](https://support.microsoft.com/office/7152b67a-8bb5-4075-ae5c-06ede5563c94) | 返回最小公倍数 |
| [LEFT、LEFTB 函数](https://support.microsoft.com/office/9203d2d2-7960-479b-84c6-1ea52b99640c) | 返回一个文本值的最左端字符 |
| [LEN、LENB 函数](https://support.microsoft.com/office/29236f94-cedc-429d-affd-b5e33d2c67cb) | 返回文本字符串中的字符数 |
| [LN 函数](https://support.microsoft.com/office/81fe1ed7-dac9-4acd-ba1d-07a142c6118f) | 返回数值的自然对数 |
| [LOG 函数](https://support.microsoft.com/office/4e82f196-1ca9-4747-8fb0-6c4a3abb3280) | 返回一个数在指定底下的对数 |
| [LOG10 函数](https://support.microsoft.com/office/c75b881b-49dd-44fb-b6f4-37e3486a0211) | 返回以 10 为底的对数 |
| [LOGNORM.DIST 函数](https://support.microsoft.com/office/eb60d00b-48a9-4217-be2b-6074aee6b070) | 返回对数正态分布 |
| [LOGNORM.INV 函数](https://support.microsoft.com/office/fe79751a-f1f2-4af8-a0a1-e151b2d4f600) | 返回对数正态分布的反函数 |
| [LOOKUP 函数](https://support.microsoft.com/office/446d94af-663b-451d-8251-369d5e3864cb) | 在向量或数组中查找值 |
| [LOWER 函数](https://support.microsoft.com/office/3f21df02-a80c-44b2-afaf-81358f9fdeb4) | 将文本转换为小写 |
| [MATCH 函数](https://support.microsoft.com/office/e8dffd45-c762-47d6-bf89-533f4a37673a) | 在引用或数组中查找值 |
| [MAX 函数](https://support.microsoft.com/office/e0012414-9ac8-4b34-9a47-73e662c08098) | 返回参数列表中的最大值 |
| [MAXA 函数](https://support.microsoft.com/office/814bda1e-3840-4bff-9365-2f59ac2ee62d) | 返回参数列表中的最大值，包括数字、文本和逻辑值 |
| [MDURATION 函数](https://support.microsoft.com/office/b3786a69-4f20-469a-94ad-33e5b90a763c) | 为假定票面值为 100 元的债券返回麦考利修正持续时间 |
| [MEDIAN 函数](https://support.microsoft.com/office/d0916313-4753-414c-8537-ce85bdd967d2) | 返回给定数字的中值 |
| [MID、MIDB 函数](https://support.microsoft.com/office/d5f9e25c-d7d6-472e-b568-4ecb12433028) | 从指定位置开始，返回文本字符串中特定数量的字符。 |
| [MIN 函数](https://support.microsoft.com/office/61635d12-920f-4ce2-a70f-96f202dcc152) | 返回参数列表中的最小值 |
| [MINA 函数](https://support.microsoft.com/office/245a6f46-7ca5-4dc7-ab49-805341bc31d3) | 返回参数列表中的最小值，包括数字、文本和逻辑值 |
| [MINUTE 函数](https://support.microsoft.com/office/af728df0-05c4-4b07-9eed-a84801a60589) | 将序列号转换为分钟 |
| [MIRR 函数](https://support.microsoft.com/office/b020f038-7492-4fb4-93c1-35c345b53524) | 返回内部收益率，它的正现金流和负现金流以不同的比率融资 |
| [MOD 函数](https://support.microsoft.com/office/9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) | 返回除法的余数 |
| [MONTH 函数](https://support.microsoft.com/office/579a2881-199b-48b2-ab90-ddba0eba86e8) | 将序列号转换为月 |
| [MROUND 函数](https://support.microsoft.com/office/c299c3b0-15a5-426d-aa4b-d2d5b3baf427) | 返回舍入到所需倍数的数值 |
| [MULTINOMIAL 函数](https://support.microsoft.com/office/6fa6373c-6533-41a2-a45e-a56db1db1bf6) | 返回一组数字的多项式 |
| [N 函数](https://support.microsoft.com/office/a624cad1-3635-4208-b54a-29733d1278c9) | 返回转换为数字的值 |
| [NA 函数](https://support.microsoft.com/office/5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c) | 返回错误值 #N/A |
| [NEGBINOM.DIST 函数](https://support.microsoft.com/office/c8239f89-c2d0-45bd-b6af-172e570f8599) | 返回负二项式分布函数值 |
| [NETWORKDAYS 函数](https://support.microsoft.com/office/48e717bf-a7a3-495f-969e-5005e3eb18e7) | 返回两个日期之间的完整工作日数 |
| [NETWORKDAYS.INTL 函数](https://support.microsoft.com/office/a9b26239-4f20-46a1-9ab8-4e925bfd5e28) | 使用能够指示哪些以及有多少天是周末的参数返回两个日期之间的完整工作日数 |
| [NOMINAL 函数](https://support.microsoft.com/office/7f1ae29b-6b92-435e-b950-ad8b190ddd2b) | 返回年度的单利 |
| [NORM.DIST 函数](https://support.microsoft.com/office/edb1cc14-a21c-4e53-839d-8082074c9f8d) | 返回正态分布函数值 |
| [NORM.INV 函数](https://support.microsoft.com/office/54b30935-fee7-493c-bedb-2278a9db7e13) | 返回正态分布的反函数 |
| [NORM.S.DIST 函数](https://support.microsoft.com/office/1e787282-3832-4520-a9ae-bd2a8d99ba88) | 返回标准正态分布函数值 |
| [NORM.S.INV 函数](https://support.microsoft.com/office/d6d556b4-ab7f-49cd-b526-5a20918452b1) | 返回标准正态分布的反函数 |
| [NOT 函数](https://support.microsoft.com/office/9cfc6011-a054-40c7-a140-cd4ba2d87d77) | 反转其参数的逻辑 |
| [NOW 函数](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46) | 返回当前日期和时间的序列号 |
| [NPER 函数](https://support.microsoft.com/office/240535b5-6653-4d2d-bfcf-b6a38151d815) | 返回一项投资的周期数量 |
| [NPV 函数](https://support.microsoft.com/office/8672cb67-2576-4d07-b67b-ac28acf2a568) | 基于一系列定期现金流和贴现率返回投资的净现值 |
| [NUMBERVALUE 函数](https://support.microsoft.com/office/1b05c8cf-2bfa-4437-af70-596c7ea7d879) | 按独立于区域设置的方式将文本转换为数字 |
| [OCT2BIN 函数](https://support.microsoft.com/office/55383471-3c56-4d27-9522-1a8ec646c589) | 将八进制数转换为二进制 |
| [OCT2DEC 函数](https://support.microsoft.com/office/87606014-cb98-44b2-8dbb-e48f8ced1554) | 将八进制数转换为十进制 |
| [OCT2HEX 函数](https://support.microsoft.com/office/912175b4-d497-41b4-a029-221f051b858f) | 将八进制数转换为十六进制 |
| [ODD 函数](https://support.microsoft.com/office/deae64eb-e08a-4c88-8b40-6d0b42575c98) | 将数值向上舍入到最接近的奇数 |
| [ODDFPRICE 函数](https://support.microsoft.com/office/d7d664a8-34df-4233-8d2b-922bcf6a69e1) | 返回每张票面为 100 元且第一期为奇数的债券的现价 |
| [ODDFYIELD 函数](https://support.microsoft.com/office/66bc8b7b-6501-4c93-9ce3-2fd16220fe37) | 返回第一期为奇数的债券的收益 |
| [ODDLPRICE 函数](https://support.microsoft.com/office/fb657749-d200-4902-afaf-ed5445027fc4) | 返回每张票面为 100 元且最后一期为奇数的债券的现价 |
| [ODDLYIELD 函数](https://support.microsoft.com/office/c873d088-cf40-435f-8d41-c8232fee9238) | 返回最后一期为奇数的债券的收益 |
| [OR 函数](https://support.microsoft.com/office/7d17ad14-8700-4281-b308-00b131e22af0) | 如果任意参数为 true，返回 `TRUE` |
| [PDURATION 函数](https://support.microsoft.com/office/44f33460-5be5-4c90-b857-22308892adaf) | 返回投资达到指定的值所需的期数 |
| [PERCENTILE.EXC 函数](https://support.microsoft.com/office/bbaa7204-e9e1-4010-85bf-c31dc5dce4ba) | 返回数组的 K 百分点值，K 介于 0 与 1 之间，不含 0 与 1 |
| [PERCENTILE.INC 函数](https://support.microsoft.com/office/680f9539-45eb-410b-9a5e-c1355e5fe2ed) | 返回数组的 K 百分点值 |
| [PERCENTRANK.EXC 函数](https://support.microsoft.com/office/d8afee96-b7e2-4a2f-8c01-8fcdedaa6314) | 返回特定数值在一个数据集中的百分比排名（介于 0 与 1 之间，不含 0 与 1） |
| [PERCENTRANK.INC 函数](https://support.microsoft.com/office/149592c9-00c0-49ba-86c1-c1f45b80463a) | 返回一组数据中的值的百分比排名 |
| [PERMUT 函数](https://support.microsoft.com/office/3bd1cb9a-2880-41ab-a197-f246a7a602d3) | 返回给定数目对象的排列数 |
| [PERMUTATIONA 函数](https://support.microsoft.com/office/6c7d7fdc-d657-44e6-aa19-2857b25cae4e) | 返回从给定元素数目的集合中选取若干（包括重复项）元素的排列数 |
| [PHI 函数](https://support.microsoft.com/office/23e49bc6-a8e8-402d-98d3-9ded87f6295c) | 返回标准正态分布的密度函数值 |
| [PI 函数](https://support.microsoft.com/office/264199d0-a3ba-46b8-975a-c4a04608989b) | 返回 pi 值 |
| [PMT 函数](https://support.microsoft.com/office/0214da64-9a63-4996-bc20-214433fa6441) | 返回年金的定期支付额 |
| [POISSON.DIST 函数](https://support.microsoft.com/office/8fe148ff-39a2-46cb-abf3-7772695d9636) | 返回泊松分布 |
| [POWER 函数](https://support.microsoft.com/office/d3f2908b-56f4-4c3f-895a-07fb519c362a) | 返回某数的乘幂结果 |
| [PPMT 函数](https://support.microsoft.com/office/c370d9e3-7749-4ca4-beea-b06c6ac95e1b) | 返回对给定期间内的投资所支付的本金 |
| [PRICE 函数](https://support.microsoft.com/office/3ea9deac-8dfa-436f-a7c8-17ea02c21b0a) | 返回每张票面为 100 元且定期支付利息的债券的现价 |
| [PRICEDISC 函数](https://support.microsoft.com/office/d06ad7c1-380e-4be7-9fd9-75e3079acfd3) | 返回每张票面为 100 元的已贴现债券的现价 |
| [PRICEMAT 函数](https://support.microsoft.com/office/52c3b4da-bc7e-476a-989f-a95f675cae77) | 返回每张票面为 100 元且在到期日支付利息的债券的现价 |
| [PRODUCT 函数](https://support.microsoft.com/office/8e6b5b24-90ee-4650-aeec-80982a0512ce) | 将其参数相乘 |
| [PROPER 函数](https://support.microsoft.com/office/52a5a283-e8b2-49be-8506-b2887b889f94) | 使一个文本值的每个词的首字母大写 |
| [PV 函数](https://support.microsoft.com/office/23879d31-0e02-4321-be01-da16e8168cbd) | 返回一项投资的当前值 |
| [QUARTILE.EXC 函数](https://support.microsoft.com/office/5a355b7a-840b-4a01-b0f1-f538c2864cad) | 基于从 0 到 1 之间（不含 0 与 1）的百分点值，返回一组数据的四分位点 |
| [QUARTILE.INC 函数](https://support.microsoft.com/office/1bbacc80-5075-42f1-aed6-47d735c4819d) | 返回一组数据的四分位点 |
| [QUOTIENT 函数](https://support.microsoft.com/office/9f7bf099-2a18-4282-8fa4-65290cc99dee) | 返回除法结果的整数部分 |
| [RADIANS 函数](https://support.microsoft.com/office/ac409508-3d48-45f5-ac02-1497c92de5bf) | 将度转换为弧度 |
| [RAND 函数](https://support.microsoft.com/office/4cbfa695-8869-4788-8d90-021ea9f5be73) | 返回 0 和 1 之间的一个随机数 |
| [RANDBETWEEN 函数](https://support.microsoft.com/office/4cc7f0d1-87dc-4eb7-987f-a469ab381685) | 返回指定数字之间的随机数 |
| [RANK.AVG 函数](https://support.microsoft.com/office/bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a) | 返回某数字在一列数字中的排名 |
| [RANK.EQ 函数](https://support.microsoft.com/office/284858ce-8ef6-450e-b662-26245be04a40) | 返回某数字在一列数字中的排名 |
| [RATE 函数](https://support.microsoft.com/office/9f665657-4a7e-4bb7-a030-83fc59e748ce) | 返回年金的定期利率 |
| [RECEIVED 函数](https://support.microsoft.com/office/7a3f8b93-6611-4f81-8576-828312c9b5e5) | 返回完全投资型债券到期收回的金额 |
| [REPLACE、REPLACEB 函数](https://support.microsoft.com/office/8d799074-2425-4a8a-84bc-82472868878a) | 替换文本中的字符 |
| [REPT 函数](https://support.microsoft.com/office/04c4d778-e712-43b4-9c15-d656582bb061) | 以给定的次数重复文本 |
| [RIGHT、RIGHTB 函数](https://support.microsoft.com/office/240267ee-9afa-4639-a02b-f19e1786cf2f) | 返回一个文本值的最右端字符 |
| [ROMAN 函数](https://support.microsoft.com/office/d6b0b99e-de46-4704-a518-b45a0f8b56f5) | 将阿拉伯数字转换为文本形式的罗马数字 |
| [ROUND 函数](https://support.microsoft.com/office/c018c5d8-40fb-4053-90b1-b3e7f61a213c) | 将数字舍入到指定位数 |
| [ROUNDDOWN 函数](https://support.microsoft.com/office/2ec94c73-241f-4b01-8c6f-17e6d7968f53) | 将数字向零的方向向下舍入 |
| [ROUNDUP 函数](https://support.microsoft.com/office/f8bc9b23-e795-47db-8703-db171d0c42a7) | 将数字向远离零的方向向上舍入 |
| [ROWS 函数](https://support.microsoft.com/office/b592593e-3fc2-47f2-bec1-bda493811597) | 返回引用中的行数 |
| [RRI 函数](https://support.microsoft.com/office/6f5822d8-7ef1-4233-944c-79e8172930f4) | 返回某项投资增长的等效利率 |
| [SEC 函数](https://support.microsoft.com/office/ff224717-9c87-4170-9b58-d069ced6d5f7) | 返回一个角度的正割值 |
| [SECH 函数](https://support.microsoft.com/office/e05a789f-5ff7-4d7f-984a-5edb9b09556f) | 返回一个角度的双曲正割值 |
| [SECOND 函数](https://support.microsoft.com/office/740d1cfc-553c-4099-b668-80eaa24e8af1) | 将序列号转换为秒 |
| [SERIESSUM 函数](https://support.microsoft.com/office/a3ab25b5-1093-4f5b-b084-96c49087f637) | 返回基于以下公式的幂级数之和 |
| [SHEET 函数](https://support.microsoft.com/office/44718b6f-8b87-47a1-a9d6-b701c06cff24) | 返回引用的工作表的工作表编号 |
| [SHEETS 函数](https://support.microsoft.com/office/770515eb-e1e8-45ce-8066-b557e5e4b80b) | 返回引用中的工作表数 |
| [SIGN 函数](https://support.microsoft.com/office/109c932d-fcdc-4023-91f1-2dd0e916a1d8) | 返回数值的符号 |
| [SIN 函数](https://support.microsoft.com/office/cf0e3432-8b9e-483c-bc55-a76651c95602) | 返回给定角的正弦值 |
| [SINH 函数](https://support.microsoft.com/office/1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7) | 返回某一数字的双曲正弦值 |
| [SKEW 函数](https://support.microsoft.com/office/bdf49d86-b1ef-4804-a046-28eaea69c9fa) | 返回一个分布的不对称度 |
| [SKEW.P 函数](https://support.microsoft.com/office/76530a5c-99b9-48a1-8392-26632d542fcb) | 基于总体返回一个分布的不对称度：用来体现某一分布相对其平均值的不对称程度 |
| [SLN 函数](https://support.microsoft.com/office/cdb666e5-c1c6-40a7-806a-e695edc2f1c8) | 返回某项资产一个周期的直线折旧值 |
| [SMALL 函数](https://support.microsoft.com/office/17da8222-7c82-42b2-961b-14c45384df07) | 返回数据集中第 k 个最小值 |
| [SQRT 函数](https://support.microsoft.com/office/654975c2-05c4-4831-9a24-2c65e4040fdf) | 返回正平方根 |
| [SQRTPI 函数](https://support.microsoft.com/office/1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4) | 返回（数字 * pi）的平方根 |
| [STANDARDIZE 函数](https://support.microsoft.com/office/81d66554-2d54-40ec-ba83-6437108ee775) | 返回正态分布概率值 |
| [STDEV.P 函数](https://support.microsoft.com/office/6e917c05-31a0-496f-ade7-4f4e7462f285) | 基于整个样本总体计算标准偏差 |
| [STDEV.S 函数](https://support.microsoft.com/office/7d69cf97-0c1f-4acf-be27-f3e83904cc23) | 基于样本估计标准偏差 |
| [STDEVA 函数](https://support.microsoft.com/office/5ff38888-7ea5-48de-9a6d-11ed73b29e9d) | 基于样本估计标准偏差，包括数字、文本和逻辑值 |
| [STDEVPA 函数](https://support.microsoft.com/office/5578d4d6-455a-4308-9991-d405afe2c28c) | 基于整个样本总体计算标准偏差，包括数字、文本和逻辑值 |
| [SUBSTITUTE 函数](https://support.microsoft.com/office/6434944e-a904-4336-a9b0-1e58df3bc332) | 在文本串中用新文本替换旧文本。 |
| [SUBTOTAL 函数](https://support.microsoft.com/office/7b027003-f060-4ade-9040-e478765b9939) | 返回一个数据列表或数据库的分类汇总 |
| [SUM 函数](https://support.microsoft.com/office/043e1c7d-7726-4e80-8f32-07b23e057f89) | 对参数求和 |
| [SUMIF 函数](https://support.microsoft.com/office/169b8c99-c05c-4483-a712-1697a653039b) | 根据给定的标准，对指定的单元格求和 |
| [SUMIFS 函数](https://support.microsoft.com/office/c9e748f5-7ea7-455d-9406-611cebce642b) | 对区域中满足多个条件的单元格求和 |
| [SUMSQ 函数](https://support.microsoft.com/office/e3313c02-51cc-4963-aae6-31442d9ec307) | 返回所有参数的平方和 |
| [SYD 函数](https://support.microsoft.com/office/069f8106-b60b-4ca2-98e0-2a0f206bdb27) | 返回某项资产在指定期间的年限总额折旧。 |
| [T 函数](https://support.microsoft.com/office/fb83aeec-45e7-4924-af95-53e073541228) | 将其参数转换为文本 |
| [T.DIST 函数](https://support.microsoft.com/office/4329459f-ae91-48c2-bba8-1ead1c6c21b2) | 返回学生 t 分布的百分点（概率） |
| [T.DIST.2T 函数](https://support.microsoft.com/office/198e9340-e360-4230-bd21-f52f22ff5c28) | 返回学生 t 分布的百分点（概率） |
| [T.DIST.RT 函数](https://support.microsoft.com/office/20a30020-86f9-4b35-af1f-7ef6ae683eda) | 返回学生的 t 分布 |
| [T.INV 函数](https://support.microsoft.com/office/2908272b-4e61-4942-9df9-a25fec9b0e2e) | 返回作为概率和自由度函数的学生 t 分布的 t 值 |
| [T.INV.2T 函数](https://support.microsoft.com/office/ce72ea19-ec6c-4be7-bed2-b9baf2264f17) | 返回学生 t 分布的反函数 |
| [TAN 函数](https://support.microsoft.com/office/08851a40-179f-4052-b789-d7f699447401) | 返回一个数字的正切值 |
| [TANH 函数](https://support.microsoft.com/office/017222f0-a0c3-4f69-9787-b3202295dc6c) | 返回一个数字的双曲正切值 |
| [TBILLEQ 函数](https://support.microsoft.com/office/2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c) | 返回短期国库券的等价债券收益 |
| [TBILLPRICE 函数](https://support.microsoft.com/office/eacca992-c29d-425a-9eb8-0513fe6035a2) | 返回每张票面为 100 元的短期国库券的现价 |
| [TBILLYIELD 函数](https://support.microsoft.com/office/6d381232-f4b0-4cd5-8e97-45b9c03468ba) | 返回短期国库券的收益 |
| [TEXT 函数](https://support.microsoft.com/office/20d5ac4d-7b94-49fd-bb38-93d29371225c) | 设置数字格式并将其转换为文本 |
| [TIME 函数](https://support.microsoft.com/office/9a5aff99-8f7d-4611-845e-747d0b8d5457) | 返回特定时间的序列号 |
| [TIMEVALUE 函数](https://support.microsoft.com/office/0b615c12-33d8-4431-bf3d-f3eb6d186645) | 将以文本表达的时间转换为序列号 |
| [TODAY 函数](https://support.microsoft.com/office/5eb3078d-a82c-4736-8930-2f51a028fdd9) | 返回当前日期的序列号 |
| [TRIM 函数](https://support.microsoft.com/office/410388fa-c5df-49c6-b16c-9e5630b479f9) | 从文本中删除空格 |
| [TRIMMEAN 函数](https://support.microsoft.com/office/d90c9878-a119-4746-88fa-63d988f511d3) | 返回数据集内部的平均值 |
| [TRUE 函数](https://support.microsoft.com/office/7652c6e3-8987-48d0-97cd-ef223246b3fb) | 返回逻辑值 `TRUE` |
| [TRUNC 函数](https://support.microsoft.com/office/8b86a64c-3127-43db-ba14-aa5ceb292721) | 将数字截断为整数 |
| [TYPE 函数](https://support.microsoft.com/office/45b4e688-4bc3-48b3-a105-ffa892995899) | 返回一个指示数值数据类型的数字 |
| [UNICHAR 函数](https://support.microsoft.com/office/ffeb64f5-f131-44c6-b332-5cd72f0659b8) | 返回给定数值引用的 Unicode 字符 |
| [UNICODE 函数](https://support.microsoft.com/office/adb74aaa-a2a5-4dde-aff6-966e4e81f16f) | 返回与文本的第一个字符相对应的数字（码位） |
| [UPPER 函数](https://support.microsoft.com/office/c11f29b3-d1a3-4537-8df6-04d0049963d6) | 将文本转换为大写 |
| [VALUE 函数](https://support.microsoft.com/office/257d0108-07dc-437d-ae1c-bc2d3953d8c2) | 将文本参数转换为数字 |
| [VAR.P 函数](https://support.microsoft.com/office/73d1285c-108c-4843-ba5d-a51f90656f3a) | 基于整个样本总体计算方差 |
| [VAR.S 函数](https://support.microsoft.com/office/913633de-136b-449d-813e-65a00b2b990b) | 基于样本估计方差 |
| [VARA 函数](https://support.microsoft.com/office/3de77469-fa3a-47b4-85fd-81758a1e1d07) | 基于样本估计方差，包括数字、文本和逻辑值 |
| [VARPA 函数](https://support.microsoft.com/office/59a62635-4e89-4fad-88ac-ce4dc0513b96) | 基于整个样本总体计算方差，包括数字、文本和逻辑值 |
| [VDB 函数](https://support.microsoft.com/office/dde4e207-f3fa-488d-91d2-66d55e861d73) | 使用余额递减法返回指定周期或部分周期内某项资产的折旧值 |
| [VLOOKUP 函数](https://support.microsoft.com/office/0bbc8083-26fe-4963-8ab8-93a18ad188a1) | 查找数组的首列并在行间移动以返回单元格的值 |
| [WEEKDAY 函数](https://support.microsoft.com/office/60e44483-2ed1-439f-8bd0-e404c190949a) | 将序列号转换为一周中的某一天 |
| [WEEKNUM 函数](https://support.microsoft.com/office/e5c43a03-b4ab-426c-b411-b18c13c75340) | 将序列号转换为代表一年中第几周的数字 |
| [WEIBULL.DIST 函数](https://support.microsoft.com/office/4e783c39-9325-49be-bbc9-a83ef82b45db) | 返回 Weibull 分布 |
| [WORKDAY 函数](https://support.microsoft.com/office/f764a5b7-05fc-4494-9486-60d494efbf33) | 返回在指定的若干个工作日之前/之后的日期（一串数字） |
| [WORKDAY.INTL 函数](https://support.microsoft.com/office/a378391c-9ba7-4678-8a39-39611a9bf81d) | 返回在指定的若干个工作日之前/之后的日期（一串数字），其中使用参数来指示哪些以及多少天为周末 |
| [XIRR 函数](https://support.microsoft.com/office/de1242ec-6477-445b-b11b-a303ad9adc9d) | 返回一组现金流的内部收益率，这些现金流不一定定期发生 |
| [XNPV 函数](https://support.microsoft.com/office/1b42bbf6-370f-4532-a0eb-d67c16b664b7) | 返回一组现金流的净现值，这些现金流不一定定期发生 |
| [XOR 函数](https://support.microsoft.com/office/1548d4c2-5e47-4f77-9a92-0533bba14f37) | 返回所有参数的逻辑“异或”值 |
| [YEAR 函数](https://support.microsoft.com/office/c64f017a-1354-490d-981f-578e8ec8d3b9) | 将序列号转换为年 |
| [YEARFRAC 函数](https://support.microsoft.com/office/3844141e-c76d-4143-82b6-208454ddc6a8) | 返回表示 start_date 和 end_date 之间的天数占一年总天数的比值 |
| [YIELD 函数](https://support.microsoft.com/office/f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe) | 返回定期支付利息的债券的收益 |
| [YIELDDISC 函数](https://support.microsoft.com/office/a9dbdbae-7dae-46de-b995-615faffaaed7) | 返回已贴现债券的年收益；例如，短期国库券 |
| [YIELDMAT 函数](https://support.microsoft.com/office/ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f) | 返回到期付息的债券的年收益 |
| [Z.TEST 函数](https://support.microsoft.com/office/d633d5a3-2031-4614-a016-92180ad82bee) | 返回 z 检验的收尾概率值 |

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [Functions 类 (JavaScript API for Excel) ](/javascript/api/excel/excel.functions)
- [工作簿函数对象 (JavaScript API Excel) ](/javascript/api/excel/excel.workbook#excel-excel-workbook-functions-member)
