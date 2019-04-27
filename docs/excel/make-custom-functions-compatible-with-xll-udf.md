---
title: 使自定义函数与 XLL 用户定义的函数兼容
description: 启用与自定义函数具有等效功能的 Excel XLL 用户定义函数的兼容性
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 09914e040c1721dd8b9e91952e5814e7a6b914e5
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356849"
---
# <a name="make-your-custom-functions-compatible-with-xll-user-defined-functions"></a>使自定义函数与 XLL 用户定义的函数兼容

如果您有现有的 Excel xll, 则可以在 Office 外接程序中构建等效的自定义函数, 以将解决方案功能扩展到其他平台 (如 online 或 macOS)。 但是, Office 外接程序没有 xll 中提供的所有功能。 根据您的解决方案使用的功能, XLL 可能比 Excel for Windows 中的 Office 外接程序自定义函数提供更好的体验。

您可以配置 Office 外接程序, 以便在用户计算机上已安装等效 XLL 时, Excel 将运行 XLL 而不是 Office 外接程序自定义函数。 xll 被称作等效操作, 因为 Excel 将在 XLL 和 Office 加载项自定义函数之间进行无缝转换, 具体取决于 Windows 上安装的功能。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>在清单中指定等效 XLL

若要启用与现有 XLL 的兼容性, 请在 Office 外接程序的清单中标识等效的 XLL。 在 Windows 上运行时, Excel 将使用 XLL 的函数而不是 Office 外接程序自定义函数。

若要设置自定义函数的等效 XLL, 请指定`FileName` XLL 的。 当用户使用 XLL 中的函数打开工作簿时, Excel 会将函数转换为兼容函数。 在 Windows Excel 中打开时, 工作簿将使用 XLL, 并且在联机或在 macOS 中打开时, 它将使用 Office 外接程序中的自定义函数。

下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。 通常, 出于完整性的考虑, 这两个示例都会在上下文中显示这两个示例。 它们`ProgID` `FileName`分别由各自标识。 有关 COM 加载项兼容性的详细信息, 请参阅[使 Office 外接程序与现有 COM 加载项兼容](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)。

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

> [!NOTE]
> 如果外接程序声明其自定义函数是 XLL 兼容的, 则稍后更改清单可能会破坏用户的工作簿, 因为它会更改文件格式。

## <a name="office-add-in-updates"></a>Office 外接程序更新

为 office 外接程序指定等效 XLL 后, Excel 将停止处理 office 外接程序的更新。 用户必须卸载 XLL 才能获取 Office 外接程序的最新更新。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 兼容函数的自定义函数行为

如果打开的电子表格中包含的 xll 函数也有等效的加载项, 则 XLL 的函数将转换为 xll 兼容的自定义函数。 在下一次保存时, 它们将在兼容模式下写入文件, 以便它们使用 XLL 和 Office 外接程序自定义函数 (当在其他平台上)。

下表比较了 XLL 用户定义函数、XLL 兼容的自定义函数和 Office 加载项自定义函数之间的功能。

|         |XLL 用户定义的函数 |XLL 兼容的自定义函数 |Office 外接自定义函数 |
|---------|---------|---------|---------|
| 支持的平台 | Windows | Windows、macOS、Excel online | Windows、macOS、Excel online |
| 支持的文件格式 | .XLSX、XLSB、XLSM、XLS | .XLSX、XLSB、XLSM | .XLSX、XLSB、XLSM |
| 公式自动完成 | 否 | 可访问 | 是 |
| 媒体 | 可通过 xlfRTD 和 XLL 回调实现。 | 是 | 是 |
| 函数的本地化 | 否 | 否。 名称和 ID 必须与现有 XLL 的函数相匹配。 | 是 |
| 可变函数 | 是 | 是 | 是 |
| 多线程重新计算支持 | 是 | 是 | 是 |
| 计算行为 | 无 UI。 在计算过程中, Excel 可能会无响应。 | 用户将看到 #BUSY! 在返回结果之前。 | 用户将看到 #BUSY! 在返回结果之前。 |
| 要求集 | 无 | 仅 customfunctions.js 1。1 | customfunctions.js 1.1 及更高版本 |

## <a name="see-also"></a>另请参阅

- [使您的 Office 外接程序与现有的 COM 外接程序兼容](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [自定义函数最佳实践](custom-functions-best-practices.md)
- [自定义函数更改日志](custom-functions-changelog.md)
- [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)