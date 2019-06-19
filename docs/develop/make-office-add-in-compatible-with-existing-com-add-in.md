---
title: 让 Office 加载项与现有 COM 加载项兼容
description: 启用 Office 加载项和等效 COM 加载项之间的兼容性
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: a18adb9841a9580d77c5110a0346f365e38e3746
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059718"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a>使您的 Office 外接程序与现有 COM 加载项兼容 (预览)

如果您有一个现有的 COM 加载项, 则可以在 Office 加载项中构建等效功能, 从而使您的解决方案能够在其他平台 (如 web 或 Mac 上的 Office) 上运行。 在某些情况下, Office 外接程序可能无法提供相应 COM 外接程序中提供的所有功能。 在这些情况下, 您的 COM 外接程序在 Windows 上提供的用户体验可能比相应的 Office 外接程序提供的更好。

您可以配置 Office 加载项, 以便在用户的计算机上已安装等效的 COM 加载项时, Windows 上的 Office 将运行 COM 加载项, 而不是 Office 外接程序。 COM 加载项称为 "等效", 因为 Office 将根据安装了用户计算机的加载项和 Office 加载项在 COM 加载项之间进行无缝转换。

> [!NOTE]
> 此功能当前处于预览阶段, 不受支持在生产环境中使用。 它在 Excel、Word 和 PowerPoint 版本16.0.11629.20214 或更高版本中可用。 若要访问此版本, 您必须拥有 Office 365 订阅, 并在**内幕**级加入[Office 预览体验成员](https://products.office.com/office-insider)计划。

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>在清单中指定等效的 COM 加载项

若要在 Office 外接程序和 COM 加载项之间启用兼容性, 请在 Office 外接程序的[清单](add-in-manifests.md)中标识等效的 COM 加载项。 然后, Windows 上的 Office 将使用 COM 加载项, 而不是 Office 加载项 (如果已安装)。

以下示例显示了将 COM 加载项指定为等效加载项的清单部分。 `ProgId`元素的值标识 COM 加载项, 并且`EquivalentAddins`元素必须紧跟在结束`VersionOverrides`标记之前。

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  <EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> 有关 COM 加载项和 XLL UDF 兼容性的信息, 请参阅[使您的自定义函数与 XLL 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。

## <a name="equivalent-behavior-for-users"></a>用户的等效行为

在 Office 外接程序清单中指定等效的 COM 外接程序时, 如果安装了等效的 COM 加载项, 则 Windows 上的 Office 将不会显示 Office 加载项的用户界面 (UI)。 Office 仅隐藏 Office 加载项的功能区按钮, 不会阻止安装。 因此, 你的 Office 外接程序仍将显示在 UI 中的以下位置:

- 在 **"我的外接程序**" 下
- 作为功能区管理器中的条目

> [!NOTE]
> 在清单中指定等效的 COM 加载项不会对其他平台 (如 web 上的 Office 或 Office for Mac) 产生影响。

以下方案描述了根据用户获取 Office 加载项的方式而发生的情况。

### <a name="appsource-acquisition-of-an-office-add-in"></a>AppSource Office 外接程序的获取

如果用户从 AppSource 获取 Office 加载项, 并且已安装等效的 COM 加载项, 则 Office 将:

1. 安装 Office 加载项。
2. 在功能区中隐藏 Office 加载项 UI。
3. 为用户显示一个指出 "COM 加载项" 功能区按钮的调用。

### <a name="centralized-deployment-of-office-add-in"></a>Office 加载项的集中部署

如果管理员使用集中部署将 Office 加载项部署到其租户, 并且已安装了等效的 COM 加载项, 则用户必须重新启动 Office 才能看到任何更改。 Office 重启后, 将执行以下操作:

1. 安装 Office 加载项。
2. 在功能区中隐藏 Office 加载项 UI。
3. 为用户显示一个指出 "COM 加载项" 功能区按钮的调用。

### <a name="document-shared-with-embedded-office-add-in"></a>与嵌入的 Office 加载项共享的文档

如果用户安装了 COM 加载项, 然后使用嵌入的 Office 外接程序获取共享文档, 然后当他们打开文档时, Office 将执行以下操作:

1. 提示用户信任 Office 加载项。
2. 如果受信任, Office 加载项将会安装。
3. 在功能区中隐藏 Office 加载项 UI。

## <a name="other-com-add-in-behavior"></a>其他 COM 加载项行为

如果用户卸载等效的 COM 加载项, 则 Windows 上的 Office 将还原 Office 加载项 UI。

为 Office 外接程序指定等效的 COM 外接程序后, Office 将停止处理 Office 外接程序的更新。 若要获取 Office 外接程序的最新更新, 用户必须先卸载 COM 加载项。

## <a name="see-also"></a>另请参阅

- [使自定义函数与 XLL 用户定义的函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)
