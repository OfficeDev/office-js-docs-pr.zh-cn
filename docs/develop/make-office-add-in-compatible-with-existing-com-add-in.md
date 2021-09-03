---
title: 确认 Office 加载项与已有的COM 加载项兼容
description: 启用加载项Office COM 加载项之间的兼容性。
ms.date: 08/03/2021
localization_priority: Normal
ms.openlocfilehash: bb842c60beb329571ce3dc7f055cc1d9d606209b
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868440"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>确认 Office 加载项与已有的COM 加载项兼容

如果你有现有的 COM 加载项，可以在 Office 加载项中生成等效功能，从而使你的解决方案可以在其他平台（如 Office web 版 或 Mac）中运行。 在某些情况下，Office加载项可能无法提供相应 COM 加载项中提供的所有功能。 在这些情况下，COM 加载项可能提供比加载项Windows更好的用户体验Office COM 加载项。

您可以配置 Office 加载项，以便当用户的计算机上已安装等效 COM 加载项时，Windows 上的 Office 将运行 COM 加载项，而不是 Office 加载项。 COM 加载项称为"等效"，因为 Office 将按照安装用户计算机时在 COM 加载项和 Office 加载项之间无缝转换。

> [!NOTE]
> 连接到订阅时，以下平台和应用程序支持Microsoft 365功能。 COM 加载项无法安装在任何其他平台上，因此在这些平台上，将忽略本文稍后讨论的清单 `EquivalentAddins` 元素。
>
> - Excel版本 1904 PowerPoint更高版本Windows (、Word 和) 
> - Outlook 2102 Windows (更高版本上的) 受支持 Exchange 服务器版本
>   - Exchange Online
>   - Exchange 2019 累积更新 10 或更高版本 ([KB5003612](https://support.microsoft.com/topic/b1434cad-3fbc-4dc3-844d-82568e8d4344)) 
>   - Exchange 2016 累积更新 21 或更高版本 ([KB5003611](https://support.microsoft.com/topic/b7ba1656-abba-4a0b-9be9-dac45095d969)) 

## <a name="specify-an-equivalent-com-add-in"></a>指定等效的 COM 加载项

### <a name="manifest"></a>清单

> [!IMPORTANT]
> 适用于 Excel、Outlook、PowerPoint 和 Word。

若要在加载项Office COM 加载项之间实现兼容性，请确定加载项清单中的等效 COM Office加载项。 [](add-in-manifests.md) 然后Office加载项Windows COM 加载项，而不是Office加载项（如果两者均已安装）。

以下示例显示清单中将 COM 加载项指定为等效加载项的部分。 元素的值标识 `ProgId` COM 加载项， [而 EquivalentAddins](../reference/manifest/equivalentaddins.md) 元素必须紧接在结束标记 `VersionOverrides` 的之前。

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
> 有关 COM 加载项和 XLL UDF 兼容性的信息，请参阅使自定义函数与 [XLL 用户定义函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)。 不适用于Outlook。

### <a name="group-policy"></a>组策略

> [!IMPORTANT]
> 仅适用于Outlook。

若要声明 Outlook Web 加载项和 COM/VSTO 加载项之间的兼容性，请标识组策略停用 **Outlook Web** 加载项中的等效 COM 加载项，这些加载项的等效 COM 或 VSTO 加载项通过配置安装在用户计算机上。 然后Outlook一Windows COM 加载项，而不是 Web 加载项（如果两者都已安装的话）。

1. 下载最新的 [管理模板工具](https://www.microsoft.com/download/details.aspx?id=49030)，注意该工具的 **安装说明**。
1. 打开 **gpedit.msc (本地组策略**) 。
1. 导航到 **"用户配置**  >  **""管理模板**   >  **""Microsoft**  >  **Outlook 2016"其他"。**
1. 选择"停用 **Outlook加载项的** 等效 COM 或 VSTO Web 加载项"设置。
1. 打开链接以编辑策略设置。
1. 在对话框中 **Outlook Web 外接程序停用：**
    1. 将 **"值** `Id` 名称"设置为在 Web 加载项清单中找到的 。 **重要** 提示 *：请勿* 在条目周围 `{}` 添加大括号。
    1. 将 **"** 值 `ProgId` "设置为等效 COM/VSTO加载项的 。
    1. 选择 **"** 确定"将更新生效。
    ![Screenshot showing the dialog "Outlook web add-ins to deactivate".](../images/outlook-deactivate-gpo-dialog.png)

## <a name="equivalent-behavior-for-users"></a>用户的等效行为

指定[等效 COM](#specify-an-equivalent-com-add-in)加载项时，如果安装了等效 COM 加载项，Windows 上的 Office 将不会显示 Office 加载项的用户界面 (UI) 。 Office隐藏加载项的功能Office按钮，不会阻止安装。 因此Office外接程序仍将显示在 UI 内的以下位置。

- 在 **"我的外接程序"下**
- 作为功能区管理器中的条目， (Excel、Word 和 PowerPoint仅) 

> [!NOTE]
> 在清单中指定等效的 COM 加载项对于其他平台（如 Office web 版 或 Mac）没有影响。

以下方案描述了根据用户如何获取加载项Office发生的情况。

### <a name="appsource-acquisition-of-an-office-add-in"></a>AppSource 获取Office加载项

如果用户从 AppSource Office加载项，并且已安装等效的 COM 加载项，Office：

1. 安装Office加载项。
2. 隐藏Office功能区中的外接程序 UI。
3. 为指出 COM 加载项功能区按钮的用户显示一个调用。

### <a name="centralized-deployment-of-office-add-in"></a>加载项Office集中部署

如果管理员使用集中式部署将 Office 外接程序部署到其租户，并且已安装等效的 COM 外接程序，则用户必须先重新启动 Office，然后才能看到任何更改。 重新启动Office，它将：

1. 安装Office加载项。
2. 隐藏Office功能区中的外接程序 UI。
3. 为指出 COM 加载项功能区按钮的用户显示一个调用。

### <a name="document-shared-with-embedded-office-add-in"></a>与嵌入加载项Office的文档

如果用户安装了 COM 加载项，然后获取与嵌入式 Office 加载项共享的文档，那么在打开该文档时，Office：

1. 提示用户信任Office外接程序。
2. 如果受信任，Office安装外接程序。
3. 隐藏Office功能区中的外接程序 UI。

## <a name="other-com-add-in-behavior"></a>其他 COM 加载项行为

### <a name="excel-powerpoint-word"></a>Excel、PowerPoint、Word

如果用户卸载等效的 COM 加载项，Office上Windows还原Office加载项 UI。

为加载项指定等效的 COM Office后，Office停止处理加载项Office更新。 若要获取加载项的最新Office，用户必须先卸载 COM 加载项。

### <a name="outlook"></a>Outlook

COM/VSTO加载项必须在启动Outlook时连接，才能禁用相应的 Web 加载项。

如果 COM/VSTO在后续 Outlook 会话期间断开连接，则 Web 外接程序在重新启动之前可能Outlook禁用。

## <a name="see-also"></a>另请参阅

- [使自定义函数与 XLL 用户定义函数兼容](../excel/make-custom-functions-compatible-with-xll-udf.md)
