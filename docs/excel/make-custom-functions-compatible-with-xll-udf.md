---
title: 使用 XLL 用户定义的函数扩展自定义函数
description: 启用与自定义函数具有等效功能的 Excel XLL 用户定义函数的兼容性
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: c34dcf5ef546fa0f337b2cbd11cca7d5e25e2de3
ms.sourcegitcommit: fecad2afa7938d7178456c11ba52b558224813b4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/09/2020
ms.locfileid: "49603776"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>使用 XLL 用户定义的函数扩展自定义函数

如果您有现有的 Excel Xll，则可以在 Excel 外接程序中构建等效的自定义函数，以将解决方案功能扩展到其他平台（如联机或 Mac）。 但是，Excel 外接程序没有在 Xll 中提供的所有功能。 根据您的解决方案使用的功能，XLL 可以提供比 excel 在 Windows 上运行的 Excel 外接程序自定义函数更好的体验。

> [!NOTE]
> 当连接到 Microsoft 365 订阅时，以下平台支持 COM 加载项和 XLL UDF 兼容性：
> - Excel 网页版
> - Windows 上的 Excel (版本1904或更高版本) 
> - Excel for Mac (版本13.329 或更高版本) 
>
> 若要在 web 上的 Excel 中使用 COM 加载项并 XLL UDF 兼容性，请使用 Microsoft 365 订阅或 [microsoft 帐户](https://account.microsoft.com/account)登录。 如果你还没有 Microsoft 365 订阅，则可以加入 [microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)，以免费的90天 renewable microsoft 365 订阅。

## <a name="specify-equivalent-xll-in-the-manifest"></a>在清单中指定等效 XLL

若要启用与现有 XLL 的兼容性，请在您的 Excel 外接程序清单中标识等效 XLL。 然后，在 Windows 上运行时，excel 将使用 XLL 的函数而不是 Excel 加载项自定义函数。

若要设置自定义函数的等效 XLL，请指定 `FileName` XLL 的。 当用户使用 XLL 中的函数打开工作簿时，Excel 会将函数转换为兼容函数。 在 Windows 上的 Excel 中打开时，工作簿将使用 XLL，并且在联机或在 Mac 上打开时，它将使用 Excel 加载项中的自定义函数。

下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。 通常会同时指定这两个。 为了实现完整性，本示例同时显示了上下文中的内容。 它们分别由各自标识 `ProgId` `FileName` 。 `EquivalentAddins`元素必须紧跟在结束 `VersionOverrides` 标记之前。 有关 COM 加载项兼容性的详细信息，请参阅 [使您的 Excel 外接程序与现有的 com 外](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)接程序兼容。

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>

    <EquivalentAddin>
      <FileName>contosofunctions.xll</FileName>
      <Type>XLL</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> 如果外接程序声明其自定义函数是 XLL 兼容的，则稍后更改清单可能会破坏用户的工作簿，因为它会更改文件格式。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 兼容函数的自定义函数行为

打开电子表格且存在等效的加载项时，外接程序的 XLL 函数将转换为 XLL 兼容的自定义函数。 在下一次保存时，XLL 函数将在兼容模式下写入文件，以便它们使用 XLL 和 Excel 外接程序自定义函数 (在其他平台上) 。

下表比较了 XLL 用户定义函数、XLL 兼容的自定义函数和 Excel 加载项自定义函数之间的功能。

|         |XLL 用户定义的函数 |XLL 兼容的自定义函数 |Excel 加载项自定义函数 |
|---------|---------|---------|---------|
| **支持的平台** | Windows | Windows、macOS、web 浏览器 | Windows、macOS、web 浏览器 |
| **支持的文件格式** | .XLSX、XLSB、XLSM、XLS | .XLSX、XLSB、XLSM | .XLSX、XLSB、XLSM |
| **公式自动完成** | 否 | 是 | 是 |
| **媒体** | 可通过 xlfRTD 和 XLL 回调实现。 | 是 | 是 |
| **函数的本地化** | 否 | 否。 名称和 ID 必须与现有 XLL 的函数相匹配。 | 是 |
| **可变函数** | 是 | 是 | 是 |
| **多线程重新计算支持** | 是 | 是 | 是 |
| **计算行为** | 无 UI。 在计算过程中，Excel 可能会无响应。 | 用户将看到 #BUSY！ 在返回结果之前。 | 用户将看到 #BUSY！ 在返回结果之前。 |
| **要求集** | 无 | Customfunctions.js 1.1 及更高版本 | Customfunctions.js 1.1 及更高版本 |

## <a name="see-also"></a>另请参阅

- [使 Excel 外接程序与现有 COM 外接程序兼容](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
