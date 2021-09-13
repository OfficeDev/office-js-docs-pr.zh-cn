---
title: 使用 Visual Studio 发布加载项
description: 如何使用 Visual Studio 2019 部署 Web 项目并打包加载项。
ms.date: 12/02/2019
ms.localizationpriority: medium
ms.openlocfilehash: 58923ff2c37edc474aefbb18fdb8ccf4fed3f079
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152555"
---
# <a name="publish-your-add-in-using-visual-studio"></a>使用 Visual Studio 发布加载项

Office 加载项包包含 XML [清单文件](../develop/add-in-manifests.md)，它可用于发布加载项。 你将不得不单独发布项目的 Web 应用程序文件。 本文介绍如何使用 Visual Studio 2019 部署 Web 项目并打包加载项。

> [!NOTE]
> 要了解如何发布使用 Yeoman 生成器创建并使用 Visual Studio Code 或任何其他编辑器开发的 Office 加载项，请参阅[发布使用 Visual Studio Code 开发的加载项](publish-add-in-vs-code.md)。

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a>使用 Visual Studio 2019 部署 Web 项目

完成以下步骤以使用 Visual Studio 2019 部署 Web 项目。

1. 从“**生成**”选项卡中，选择“**发布 [加载项名称]**”。

2. 在“**选取发布目标**”窗口中，选择其中一个选项以发布到你的首选目标。 每个发布目标都要求你提供有关入门的详细信息，例如 Azure 虚拟机或文件夹位置。 指定发布位置并填写所有必需信息后，选择“**发布**”

    > [!NOTE]
    > 选取发布目标可指定要部署到的服务器、登录服务器所需的凭据、要部署的数据库以及其他部署选项。

3. 有关每个发布目标选项的部署步骤的详细信息，请参阅[初探 Visual Studio 中的部署](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019&preserve-view=true)。

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a>使用 Visual Studio 2019 通过 IIS、FTP 或 Web 部署方法打包并发布加载项

完成以下步骤以使用 Visual Studio 2019 打包加载项。

1. 从“**生成**”选项卡中，选择“**发布 [加载项名称]**”。
2. 在“**选取发布目标**”窗口中，选择“**IIS、FTP 等**”，然后选择“**配置**”。 接下来，选择“**发布**”。
3. 此时将显示一个向导，它将指导你完成该过程。 确保发布方法是你的首选方法，例如 Web 部署。
4. 在“**目标 URL**”框中，输入托管加载项内容文件的网站的 URL，然后选择“**下一步**”。 如果计划将加载项提交到 AppSource，可以选择“**验证连接**”按钮，以发现任何可能会导致加载项遭拒的问题。 应先解决所有问题，再将加载项提交到 Microsoft Store。
5. 确认所需的任何设置（包括“**文件发布选项**”），然后选择“**保存**”。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure 网站自动提供 HTTPS 终结点。

现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。XML 清单位于 `app.publish` 文件夹的 `OfficeAppManifests` 中。例如：

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a>另请参阅

- [发布 Office 加载项](../publish/publish.md)
- [将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-the-office-store)
