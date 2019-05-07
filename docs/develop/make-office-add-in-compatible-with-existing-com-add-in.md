---
title: 使 Excel 外接程序与现有 COM 外接程序兼容
description: 启用与与 Excel 外接程序具有相同功能的等效 COM 加载项的兼容性
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 0890e14466a2cd8f5aff2d1bcf307a43cff28127
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628170"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a>使您的 Office 外接程序与现有 COM 加载项兼容 (预览)

如果您有一个现有的 COM 加载项, 则可以在 Excel 加载项中构建等效的功能, 以将解决方案功能扩展到其他平台 (如 online 或 macOS)。 但是, Excel 外接程序没有在 COM 加载项中提供的所有功能。你的 COM 加载项可以提供比 Windows 上的 Excel 外接程序更好的体验。

您可以配置 Excel 加载项, 以便在用户的计算机上已安装等效的 COM 加载项时, Office 将运行 COM 加载项, 而不是 Excel 外接程序。 COM 加载项称为 "等效", 因为 Office 将根据 Windows 上安装的 COM 加载项和 Excel 加载项之间无缝转换。

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>在清单中指定等效的 COM 加载项

若要启用与现有 COM 加载项的兼容性, 请在 Excel 外接程序的清单中标识等效的 COM 加载项。 然后, 在 Windows 上运行时, Office 将使用 COM 加载项, 而不是 Excel 外接程序。

`ProgID`指定等效 COM 加载项的。 在安装 COM 加载项时, Office 将使用 COM 加载项 UI, 而不是 Excel 外接程序的 UI。

下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。 通常, 出于完整性的考虑, 这两个示例都会在上下文中显示这两个示例。 它们`ProgID` `FileName`分别由各自标识。 有关 XLL 兼容性的详细信息, 请参阅[使您的自定义函数与 xll 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。

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

## <a name="equivalent-behavior-for-users"></a>用户的等效行为

当在 Excel 外接程序清单中指定了等效的 COM 加载项时, Office 将在安装等效 COM 加载项时禁止在 Windows 上使用 Excel 外接程序的 UI。 这不会影响其他平台 (如 online 或 macOS) 上的 Excel 外接程序 UI。 Office 仅隐藏功能区按钮, 不会阻止安装。 因此, Excel 外接程序仍将显示在以下 UI 位置:

- 在 **"我的外接程序**" 下, 因为它已安装技术。
- 作为功能区管理器中的条目。

以下方案描述了根据用户获取 Excel 加载项的方式而发生的情况。

### <a name="appsource-acquisition-of-an-excel-add-in"></a>AppSource 获取 Excel 外接程序

如果用户从 AppSource 下载 Excel 加载项, 并且已安装等效的 COM 加载项, 则 Office 将执行以下操作:

1. 安装 Excel 加载项。
2. 在功能区中隐藏 Excel 加载项 UI。
3. 为用户显示一个指出 "COM 加载项" 功能区按钮的调用。

### <a name="centralized-deployment-of-excel-add-in"></a>Excel 加载项的集中部署

如果管理员使用集中部署将 Excel 加载项部署到其租户, 并且已安装等效的 COM 加载项, 则用户需要先重新启动 Office, 然后他们才会看到任何更改。 Office 重启后, 将执行以下操作:

1. 安装 Excel 加载项。
2. 在功能区中隐藏 Excel 加载项 UI。
3. 为用户显示一个指出 "COM 加载项" 功能区按钮的调用。

### <a name="document-shared-with-embedded-excel-add-in"></a>与嵌入的 Excel 加载项共享的文档

如果用户安装了 COM 外接程序, 然后使用嵌入的 Excel 加载项获取共享文档, 然后当他们打开文档时, Office 将执行以下操作:

1. 提示用户信任 Excel 加载项。
2. 如果受信任, 将安装 Excel 加载项。
3. 在功能区中隐藏 Excel 加载项 UI。

## <a name="other-com-add-in-behavior"></a>其他 COM 加载项行为

如果用户卸载 COM 加载项, 则 Office 将在 Windows 上还原 Excel 外接程序 UI, 以获取等效的已安装 Excel 加载项。

为 Excel 加载项指定等效的 COM 加载项后, Office 将停止处理 Excel 加载项的更新。 用户必须卸载 COM 加载项, 才能获取 Excel 外接程序的最新更新。

## <a name="see-also"></a>另请参阅

- [使自定义函数与 XLL 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)
