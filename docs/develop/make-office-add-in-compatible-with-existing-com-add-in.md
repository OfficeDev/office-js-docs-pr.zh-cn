---
title: 让 Office 加载项与现有 COM 加载项兼容
description: 启用 Office 加载项和等效 COM 加载项之间的兼容性
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: ff47b75e8e560bc891c84dc839b7eceffb2400be
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609420"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>让 Office 加载项与现有 COM 加载项兼容

如果您有一个现有的 COM 加载项，则可以在 Office 加载项中构建等效功能，从而使您的解决方案能够在其他平台（如 web 或 Mac 上的 Office）上运行。 在某些情况下，Office 外接程序可能无法提供相应 COM 外接程序中提供的所有功能。 在这些情况下，您的 COM 外接程序在 Windows 上提供的用户体验可能比相应的 Office 外接程序提供的更好。

您可以配置 Office 加载项，以便在用户的计算机上已安装等效的 COM 加载项时，Windows 上的 Office 将运行 COM 加载项，而不是 Office 外接程序。 COM 加载项称为 "等效"，因为 Office 将根据安装了用户计算机的加载项和 Office 加载项在 COM 加载项之间进行无缝转换。

> [!NOTE]
> 当连接到 Office 365 订阅时，以下平台支持此功能：
> - 网页上的 Excel、Word 和 PowerPoint
> - Windows 上的 Excel、Word 和 PowerPoint （版本1904或更高版本）
> - Excel、Word 和 PowerPoint on Mac （版本13.329 或更高版本）

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>在清单中指定等效的 COM 加载项

若要在 Office 外接程序和 COM 加载项之间启用兼容性，请在 Office 外接程序的[清单](add-in-manifests.md)中标识等效的 COM 加载项。 然后，Windows 上的 Office 将使用 COM 加载项，而不是 Office 加载项（如果已安装）。

以下示例显示了将 COM 加载项指定为等效加载项的清单部分。 元素的值 `ProgId` 标识 COM 加载项，并且 `EquivalentAddins` 元素必须紧跟在结束 `VersionOverrides` 标记之前。

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> 有关 COM 加载项和 XLL UDF 兼容性的信息，请参阅[使您的自定义函数与 XLL 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。

## <a name="equivalent-behavior-for-users"></a>用户的等效行为

在 Office 外接程序清单中指定等效的 COM 外接程序时，如果安装了等效的 COM 加载项，则 Windows 上的 Office 将不会显示 Office 加载项的用户界面（UI）。 Office 仅隐藏 Office 加载项的功能区按钮，不会阻止安装。 因此，你的 Office 外接程序仍将显示在 UI 中的以下位置：

- 在 **"我的外接程序**" 下
- 作为功能区管理器中的条目

> [!NOTE]
> 在清单中指定等效的 COM 加载项不会对 web 或 Mac 等其他平台（如 Office）产生影响。

以下方案描述了根据用户获取 Office 加载项的方式而发生的情况。

### <a name="appsource-acquisition-of-an-office-add-in"></a>AppSource Office 外接程序的获取

如果用户从 AppSource 获取 Office 加载项，并且已安装等效的 COM 加载项，则 Office 将：

1. 安装 Office 加载项。
2. 在功能区中隐藏 Office 加载项 UI。
3. 为用户显示一个指出 "COM 加载项" 功能区按钮的调用。

### <a name="centralized-deployment-of-office-add-in"></a>Office 加载项的集中部署

如果管理员使用集中部署将 Office 加载项部署到其租户，并且已安装了等效的 COM 加载项，则用户必须重新启动 Office 才能看到任何更改。 Office 重启后，将执行以下操作：

1. 安装 Office 加载项。
2. 在功能区中隐藏 Office 加载项 UI。
3. 为用户显示一个指出 "COM 加载项" 功能区按钮的调用。

### <a name="document-shared-with-embedded-office-add-in"></a>与嵌入的 Office 加载项共享的文档

如果用户安装了 COM 加载项，然后使用嵌入的 Office 外接程序获取共享文档，然后当他们打开文档时，Office 将执行以下操作：

1. 提示用户信任 Office 加载项。
2. 如果受信任，Office 加载项将会安装。
3. 在功能区中隐藏 Office 加载项 UI。

## <a name="other-com-add-in-behavior"></a>其他 COM 加载项行为

如果用户卸载等效的 COM 加载项，则 Windows 上的 Office 将还原 Office 加载项 UI。

为 Office 外接程序指定等效的 COM 外接程序后，Office 将停止处理 Office 外接程序的更新。 若要获取 Office 外接程序的最新更新，用户必须先卸载 COM 加载项。

## <a name="see-also"></a>另请参阅

- [使自定义函数与 XLL 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)
