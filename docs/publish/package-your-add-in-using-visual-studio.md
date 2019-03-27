---
title: 使用 Visual Studio 打包加载项以准备发布 | Microsoft Docs
description: 如何使用 Visual Studio 2017 部署 Web 项目并打包加载项。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 9233ebed217c9e4cc5def0dace67043f29462296
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871260"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>使用 Visual Studio 打包加载项以准备发布

Office 加载项包包含 XML [清单文件](../develop/add-in-manifests.md)，它可用于发布加载项。 你将不得不单独发布项目的 Web 应用程序文件。 本文介绍如何使用 Visual Studio 2017 部署 Web 项目并打包加载项。

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a>使用 Visual Studio 2017 部署 Web 项目

完成以下步骤以使用 Visual Studio 2017 部署 Web 项目。

1. 在“**解决方案资源管理器**”中，打开外接程序项目的快捷菜单，然后选择“**发布**”。

    将显示“**发布外接程序**”页。

2. 选择“当前配置文件”**** 下拉列表中的配置文件，或选择“新建…”**** 新建配置文件。

    > [!NOTE]
    > 发布配置文件指定你要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。

    如果你选择“**新建...**”，则向导将会显示“**创建发布配置文件**”页。 可以使用此向导从托管提供程序（如 Microsoft Azure）的网站导入发布配置文件，或创建新配置文件并添加你的服务器、凭据以及下一过程中的其他设置。

    有关导入发布配置文件或创建新发布配置文件的详细信息，请参阅[创建发布配置文件](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)。

3. 在“**发布加载项**”页中，选择“**部署 Web 项目**”链接。

    将显示“**发布**”对话框。 有关如何使用此向导的详细信息，请参阅[操作方法：使用 Visual Studio 中的一键式发布来部署 Web 项目](https://msdn.microsoft.com/library/dd465337.aspx)。

## <a name="to-package-your-add-in-using-visual-studio-2017"></a>使用 Visual Studio 2017 打包加载项

完成以下步骤以使用 Visual Studio 2017 打包加载项。

1. 在“**发布加载项**”页上，选择“**打包加载项**”按钮。

    此时向导将显示“**打包加载项**”页面。

2. 在“你的网站托管在哪里?”**** 下拉列表中，选择或输入托管加载项内容文件的网站的 URL，然后选择“完成”****。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure 网站自动提供 HTTPS 终结点。

    此时，Visual Studio 生成发布加载项所需的文件，并打开发布输出文件夹。

如果计划将加载项提交到 AppSource，可以选择“**执行验证检查**”按钮，以发现任何可能会导致加载项遭拒的问题。 应先解决所有问题，再将加载项提交到 Microsoft Store。

现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。XML 清单位于 `app.publish` 文件夹的 `OfficeAppManifests` 中。例如：

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a>另请参阅

- [发布 Office 加载项](../publish/publish.md)
- [将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-the-office-store)
