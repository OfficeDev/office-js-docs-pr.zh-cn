---
title: 使用 XLL 用户定义函数扩展自定义函数
description: 启用与Excel等效功能的 XLL 用户定义函数的兼容性
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 82d1120e68a69bee74a6fe1911bbd8d3ccb3fb00
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990710"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>使用 XLL 用户定义函数扩展自定义函数

> [!NOTE]
> XLL 加载项是一个Excel扩展名 **.xll 的加载项文件**。 XLL 文件是动态链接库的一 (DLL) 文件，它只能由 Excel。 XLL 加载项文件必须使用 C 或 C++ 编写。 若要[了解Excel，请参阅开发 XLL。](/office/client-developer/excel/developing-excel-xlls)

如果你有现有的 Excel XLL 加载项，可以使用 Excel JavaScript API 生成等效的自定义函数加载项，以将解决方案功能扩展到其他平台（如 Excel web 版 或 Mac 上）。 但是Excel JavaScript API 加载项并不具有 XLL 加载项中提供的所有功能。根据解决方案使用的功能，XLL 加载项可能会提供比 Windows 上 Excel JavaScript API Excel更好的体验。

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>在清单中指定等效的 XLL

若要启用与现有 XLL 加载项的兼容性，在 JavaScript API 加载项清单中Excel等效的 XLL 加载项。 Excel在 Windows 上运行时，Excel使用 XLL 加载项函数，而不是 Excel JavaScript API Windows。

若要为自定义函数设置等效的 XLL 加载项，请指定 `FileName` XLL 文件的 。 当用户打开包含 XLL 文件中函数的工作簿时，Excel函数转换为兼容函数。 然后，在 Windows 上的 Excel 中打开工作簿时，工作簿将使用 XLL 文件，当在 Web 或 Mac 上打开时，它将使用 Excel JavaScript API 加载项中的自定义函数。

以下示例演示如何将 COM 加载项和 XLL 加载项指定为 Excel JavaScript API 加载项清单文件中等效项。 通常，您将同时指定这两者。 为完整，此示例在上下文中显示这两者。 它们分别由它们 `ProgId` 和 `FileName` 标识。 `EquivalentAddins`元素必须紧接在结束标记 `VersionOverrides` 的之前。 有关 COM 加载项兼容性的详细信息，请参阅使Office[加载项与现有 COM](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)加载项兼容。

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
> 如果 Excel JavaScript API 加载项声明其自定义函数与 XLL 加载项兼容，以后更改清单可能会破坏用户的工作簿，因为它将更改文件格式。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 兼容函数的自定义函数行为

打开电子表格并且有等效的加载项可用时，加载项的 XLL 函数将转换为 XLL 兼容的自定义函数。 下一次保存时，XLL 函数会以兼容模式写入文件，以便它们在其他平台) 中同时使用 XLL 加载项和 Excel JavaScript API 加载项自定义函数 (。

下表比较了 XLL 用户定义函数、XLL 兼容自定义函数Excel JavaScript API 加载项自定义函数之间的功能。

|         |XLL 用户定义函数 |XLL 兼容的自定义函数 |ExcelJavaScript API 加载项自定义函数 |
|---------|---------|---------|---------|
| **支持的平台** | Windows | Windows、macOS、Web 浏览器 | Windows、macOS、Web 浏览器 |
| **支持的文件格式** | XLSX、XLSB、XLSM、XLS | XLSX、XLSB、XLSM | XLSX、XLSB、XLSM |
| **公式自动完成** | 否 | 是 | 是 |
| **流式** | 可通过 xlfRTD 和 XLL 回调实现。 | 是 | 是 |
| **函数本地化** | 否 | 否。 Name 和 ID 必须与现有的 XLL 函数匹配。 | 是 |
| **可变函数** | 是 | 是 | 是 |
| **多线程重新计算支持** | 是 | 是 | 是 |
| **计算行为** | 无 UI。 Excel计算期间可能无响应。 | 用户将看到#BUSY！ 直到返回结果。 | 用户将看到#BUSY！ 直到返回结果。 |
| **要求集** | 不适用 | CustomFunctions 1.1 及更高版本 | CustomFunctions 1.1 及更高版本 |

## <a name="see-also"></a>另请参阅

- [让 Office 加载项与现有 COM 加载项兼容](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
