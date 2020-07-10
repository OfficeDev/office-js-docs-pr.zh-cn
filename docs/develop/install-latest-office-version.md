---
title: 安装最新版本 Office
description: 与如何选择获取最新版 Office 相关的信息。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: df10d64d69b64283321bbad79aca7f7f6d482dd1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093614"
---
# <a name="install-the-latest-version-of-office"></a>安装最新版本 Office

新开发人员功能（包括仍处于预览阶段的功能）会先向选择获取最新版 Office 的订阅者提供。

## <a name="opt-in-to-getting-the-latest-builds"></a>选择获取最新版

若要选择获取最新版 Office，请执行以下操作：

- 如果你是 Microsoft 365 家庭版、个人版或大学订阅者，请参阅[成为 Office 预览体验成员](https://insider.office.com)。
- 如果你是 Microsoft 365 的商业客户相关应用程序，请参阅[为商业客户的 microsoft 365 应用程序安装首次发布](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)版本。
- 如果在 Mac 上运行 Office：
  - 启动 Office 应用程序。
  - 选择“帮助”菜单上的“检查更新”****。
  - 选中“Microsoft 自动更新”框，以加入 Office 预览体验成员计划。

## <a name="get-the-latest-build"></a>获取最新版

若要获取最新版 Office，请执行以下操作：

1. 下载 [Office 部署工具](https://www.microsoft.com/download/details.aspx?id=49117)。
2. Run the tool. This extracts the following two files: Setup.exe and configuration.xml.
3. 将 configuration.xml 文件替换为[首次发布配置文件](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)。
4. 以管理员身份运行下面的命令：`setup.exe /configure configuration.xml`

> [!NOTE]
> 此命令可能需要运行很长时间才能完成，而且不会显示进度。

在安装进程完成后，你已安装最新的 Office 应用程序。 要验证你是否拥有最新版本，请从任何 Office 应用程序转到“文件”**** > “帐户”****。 在“Office 更新”下，你将看到版本号上面的 (Office Insiders) 标签。

![显示产品信息的屏幕截图（带有 Office Insiders 标签）](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Office JavaScript API 要求集对应的最低 Office 内部版本

若要了解 API 要求集对应的各个平台的最低产品内部版本，请参阅以下资源：

- [Excel JavaScript API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md)
- [OneNote JavaScript API 要求集](../reference/requirement-sets/onenote-api-requirement-sets.md)
- [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [PowerPoint JavaScript API 要求集](../reference/requirement-sets/powerpoint-api-requirement-sets.md)
- [Word JavaScript API 要求集](../reference/requirement-sets/word-api-requirement-sets.md)
- [对话框 API 要求集](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Office 通用 API 要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)
