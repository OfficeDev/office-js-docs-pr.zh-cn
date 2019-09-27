---
title: 使用 XLL 用户定义的函数扩展自定义函数
description: 启用与自定义函数具有等效功能的 Excel XLL 用户定义函数的兼容性
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: 7ec853e5b4d03267e1c9d33d2df8a79d86860095
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235300"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>使用 XLL 用户定义的函数扩展自定义函数

如果您有现有的 Excel Xll，则可以在 Excel 外接程序中构建等效的自定义函数，以将解决方案功能扩展到其他平台（如 online 或 macOS）。 但是，Excel 外接程序没有在 Xll 中提供的所有功能。 根据您的解决方案使用的功能，XLL 可以提供比 excel 在 Windows 上运行的 Excel 外接程序自定义函数更好的体验。

> [!NOTE]
> 当连接到 Office 365 订阅时，以下平台支持 COM 加载项和 XLL UDF 兼容性：
> - 在 web 上的 Excel
> - Windows 上的 Excel （版本1904或更高版本）
> - Excel for Mac （版本13.329 或更高版本）
> 
> 若要在 web 上的 Excel 中使用 COM 加载项并 XLL UDF 兼容性，请使用 Office 365 订阅或[Microsoft 帐户](https://account.microsoft.com/account)登录。 如果还没有 Office 365 订阅，可以通过加入 [Office 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个订阅。

## <a name="specify-equivalent-xll-in-the-manifest"></a>在清单中指定等效 XLL

若要启用与现有 XLL 的兼容性，请在您的 Excel 外接程序清单中标识等效 XLL。 在 Windows 上运行时，Excel 将使用 XLL 的函数而不是 Excel 加载项自定义函数。

若要设置自定义函数的等效 XLL，请指定`FileName` XLL 的。 当用户使用 XLL 中的函数打开工作簿时，Excel 会将函数转换为兼容函数。 在 Windows 上的 Excel 中打开时，工作簿将使用 XLL，并且在联机或在 macOS 中打开时，它将使用 Excel 外接程序中的自定义函数。

下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。 通常，出于完整性的考虑，这两个示例都会在上下文中显示这两个示例。 它们`ProgId` `FileName`分别由各自标识。 `EquivalentAddins`元素必须紧跟在结束`VersionOverrides`标记之前。 有关 COM 加载项兼容性的详细信息，请参阅[使您的 Excel 外接程序与现有的 com 外](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)接程序兼容。

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
  <EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> 如果外接程序声明其自定义函数是 XLL 兼容的，则稍后更改清单可能会破坏用户的工作簿，因为它会更改文件格式。

## <a name="excel-add-in-updates"></a>Excel 加载项更新

为 Excel 加载项指定等效 XLL 后，Excel 将停止处理 Excel 加载项的更新。 用户必须卸载 XLL 才能获取 Excel 外接程序的最新更新。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 兼容函数的自定义函数行为

如果打开的电子表格中包含的 XLL 函数也有等效的加载项，则 XLL 的函数将转换为 XLL 兼容的自定义函数。 在下一次保存时，它们将在兼容模式下写入文件，以便它们使用 XLL 和 Excel 外接程序自定义函数（当在其他平台上）。

下表比较了 XLL 用户定义函数、XLL 兼容的自定义函数和 Excel 加载项自定义函数之间的功能。

|         |XLL 用户定义的函数 |XLL 兼容的自定义函数 |Excel 加载项自定义函数 |
|---------|---------|---------|---------|
| 支持的平台 | Windows | Windows、macOS、Excel 网页 | Windows、macOS、Excel 网页 |
| 支持的文件格式 | .XLSX、XLSB、XLSM、XLS | .XLSX、XLSB、XLSM | .XLSX、XLSB、XLSM |
| 公式自动完成 | 否 | 可访问 | 是 |
| 媒体 | 可通过 xlfRTD 和 XLL 回调实现。 | 否 | 可访问 |
| 函数的本地化 | 否 | 否。 名称和 ID 必须与现有 XLL 的函数相匹配。 | 是 |
| 可变函数 | 是 | 是 | 是 |
| 多线程重新计算支持 | 是 | 是 | 是 |
| 计算行为 | 无 UI。 在计算过程中，Excel 可能会无响应。 | 用户将看到 #BUSY！ 在返回结果之前。 | 用户将看到 #BUSY！ 在返回结果之前。 |
| 要求集 | 不适用 | Customfunctions.js 1.1 及更高版本 | Customfunctions.js 1.1 及更高版本 |

## <a name="see-also"></a>另请参阅

- [使 Excel 外接程序与现有 COM 外接程序兼容](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
