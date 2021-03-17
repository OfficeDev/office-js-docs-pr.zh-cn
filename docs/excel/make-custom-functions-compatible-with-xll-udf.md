---
title: 使用 XLL 用户定义函数扩展自定义函数
description: 启用与 Excel XLL 用户定义函数的兼容性，这些函数具有与自定义函数等效的功能
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 32146e7eebb963e8d800b619ef052457e40f2ac6
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836814"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>使用 XLL 用户定义函数扩展自定义函数

如果已有 Excel XLL，可以在 Excel 加载项中生成等效自定义函数，以将解决方案功能扩展到其他平台（如联机平台或 Mac 平台）。 但是，Excel 加载项并不具有 XLL 中提供的所有功能。 根据您的解决方案使用的功能，XLL 可能会提供比 Windows 上的 Excel 中的 Excel 外接程序自定义函数更好的体验。

> [!NOTE]
> 连接到 Microsoft 365 订阅时，以下平台支持 COM 加载项和 XLL UDF 兼容性：
>
> - Excel 网页版
> - Windows 版 Excel (版本 1904 或更高版本) 
> - Mac 版 Excel (版本 13.329 或更高版本) 
>
> 若要在 Excel 网页版内使用 COM 加载项和 XLL UDF 兼容性，请使用 Microsoft 365 订阅或 Microsoft 帐户 [登录](https://account.microsoft.com/account)。 如果你还没有 Microsoft 365 订阅，则可以通过加入 Microsoft 365 开发人员计划获得为期 90 天的免费可续订 [Microsoft 365 订阅](https://developer.microsoft.com/office/dev-program)。

## <a name="specify-equivalent-xll-in-the-manifest"></a>在清单中指定等效的 XLL

若要启用与现有 XLL 的兼容性，请确定 Excel 加载项清单中的等效 XLL。 然后，当在 Windows 上运行时，Excel 将使用 XLL 函数，而不是 Excel 加载项自定义函数。

若要为自定义函数设置等效的 XLL，请 `FileName` 指定 XLL 的 。 当用户使用 XLL 中的函数打开工作簿时，Excel 会将函数转换为兼容函数。 然后，在 Windows 上的 Excel 中打开工作簿时，工作簿将使用 XLL，当联机打开或在 Mac 上打开时，它将使用 Excel 加载项中的自定义函数。

以下示例演示如何将 COM 加载项和 XLL 指定为等效项。 通常，您将同时指定这两者。 为完整，此示例在上下文中显示这两者。 它们分别由它们 `ProgId` 和 `FileName` 标识。 `EquivalentAddins`元素必须紧接在结束标记 `VersionOverrides` 的之前。 有关 COM 加载项兼容性的详细信息，请参阅使 [Office](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)加载项与现有 COM 加载项兼容。

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
> 如果加载项声明其自定义函数与 XLL 兼容，以后更改清单可能会破坏用户的工作簿，因为它将更改文件格式。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 兼容函数的自定义函数行为

打开电子表格且有等效的加载项可用时，加载项的 XLL 函数将转换为 XLL 兼容的自定义函数。 下一次保存时，XLL 函数会以兼容模式写入文件，以便它们在其他平台上使用 XLL 和 Excel (自定义函数) 。

下表比较了 XLL 用户定义函数、XLL 兼容自定义函数和 Excel 加载项自定义函数之间的功能。

|         |XLL 用户定义函数 |XLL 兼容的自定义函数 |Excel 加载项自定义函数 |
|---------|---------|---------|---------|
| **支持的平台** | Windows | Windows、macOS、Web 浏览器 | Windows、macOS、Web 浏览器 |
| **支持的文件格式** | XLSX、XLSB、XLSM、XLS | XLSX、XLSB、XLSM | XLSX、XLSB、XLSM |
| **公式自动完成** | 否 | 是 | 是 |
| **流式** | 可通过 xlfRTD 和 XLL 回调实现。 | 是 | 是 |
| **函数本地化** | 否 | 不正确。 Name 和 ID 必须与现有的 XLL 函数匹配。 | 是 |
| **可变函数** | 是 | 是 | 是 |
| **多线程重新计算支持** | 是 | 是 | 是 |
| **计算行为** | 无 UI。 Excel 在计算过程中可能无响应。 | 用户将看到#BUSY！ 直到返回结果。 | 用户将看到#BUSY！ 直到返回结果。 |
| **要求集** | 不适用 | CustomFunctions 1.1 及更高版本 | CustomFunctions 1.1 及更高版本 |

## <a name="see-also"></a>另请参阅

- [让 Office 加载项与现有 COM 加载项兼容](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
