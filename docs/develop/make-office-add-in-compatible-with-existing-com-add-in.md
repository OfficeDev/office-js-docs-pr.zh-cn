---
title: 使您的 Office 外接程序与现有的 COM 外接程序兼容
description: 启用与与 Office 外接程序具有相同功能的等效 COM 加载项的兼容性
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356847"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>使您的 Office 外接程序与现有的 COM 外接程序兼容

如果有现有的 COM 加载项, 则可以在 Office 外接程序中生成等效功能, 以将解决方案功能扩展到其他平台 (如 online 或 macOS)。 但是, Office 外接程序没有在 COM 加载项中提供的所有功能。在 Excel、Word 和 PowerPoint 中, 您的 COM 加载项可以提供比 Windows 上的 Office 外接程序更好的体验。

您可以配置 office 加载项, 以便在用户的计算机上已安装等效的 COM 加载项时, office 将运行 COM 加载项, 而不是 office 外接程序。 com 加载项称为 "等效", 因为 Office 将在 COM 加载项和 Office 加载项之间进行无缝转换, 具体取决于 Windows 上安装的版本。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>在清单中指定等效的 COM 加载项

若要启用与现有 COM 加载项的兼容性, 请在 Office 外接程序的清单中标识等效的 COM 加载项。 在 Windows 上运行时, office 将使用 COM 加载项, 而不是 office 外接程序。

`ProgID`指定等效 COM 加载项的。 然后, 在安装 com 加载项时, office 将使用 com 加载项 ui, 而不是 office 外接程序的 ui。

下面的示例演示如何将 COM 外接程序和 XLL 都指定为等效项。 通常, 出于完整性的考虑, 这两个示例都会在上下文中显示这两个示例。 它们`ProgID` `FileName`分别由各自标识。 有关 xll 兼容性的详细信息, 请参阅[使您的自定义函数与 xll 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。

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

当 office 外接程序清单中指定了等效的 com 加载项时, office 将在安装等效 com 加载项时在 Windows 上取消使用 office 外接程序的 UI。 这不会影响其他平台 (如 online 或 macOS) 上的 Office 外接程序的 UI。 Office 仅隐藏功能区按钮, 不会阻止安装。 因此, 你的 Office 外接程序仍将显示在以下 UI 位置:

- 在 **"我的外接程序**" 下, 因为它已安装技术。
- 作为功能区管理器中的条目。

以下方案描述了根据用户获取 Office 加载项的方式而发生的情况。

### <a name="appsource-acquisition-of-an-office-add-in"></a>AppSource Office 外接程序的获取

如果用户从 AppSource 下载 Office 加载项, 并且已安装了等效的 COM 加载项, 则 Office 将执行以下操作:

1. 安装 Office 加载项。
2. 在功能区中隐藏 Office 加载项 UI。
3. 为用户显示一个指出 "COM 加载项" 功能区按钮的调用。

### <a name="centralized-deployment-of-office-add-in"></a>Office 加载项的集中部署

如果管理员使用集中部署将 office 外接程序部署到其租户, 并且已安装等效的 COM 加载项, 则用户需要先重新启动 office, 然后他们才会看到任何更改。 Office 重启后, 将执行以下操作:

1. 安装 Office 加载项。
2. 在功能区中隐藏 Office 加载项 UI。
3. 为用户显示一个指出 "COM 加载项" 功能区按钮的调用。

### <a name="document-shared-with-embedded-office-add-in"></a>与嵌入的 Office 加载项共享的文档

如果用户安装了 COM 加载项, 然后使用嵌入的 Office 外接程序获取共享文档, 然后当他们打开文档时, Office 将执行以下操作:

1. 提示用户信任 Office 加载项。
2. 如果受信任, Office 加载项将会安装。
3. 在功能区中隐藏 Office 加载项 UI。

## <a name="other-com-add-in-behavior"></a>其他 COM 加载项行为

如果用户卸载 COM 加载项, 则 office 将在 Windows 上还原 office 外接程序 UI, 以获取等效的已安装 office 外接程序。

为 office 外接程序指定等效 COM 外接程序后, office 将停止处理 office 外接程序的更新。 用户必须卸载 COM 加载项, 才能获取 Office 外接程序的最新更新。

## <a name="see-also"></a>另请参阅

- [使自定义函数与 XLL 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)
