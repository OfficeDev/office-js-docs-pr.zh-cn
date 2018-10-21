---
title: 使用 Visual Studio 打包加载项以准备发布 | Microsoft Docs
description: 如何使用 Visual Studio 2017 部署 Web 项目并打包加载项。
ms.date: 01/25/2018
ms.openlocfilehash: 3515f88e41bc5f0af62a3b043beae5177f3291ac
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681761"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>使用 Visual Studio 打包加载项以准备发布

Office 加载项包包含 XML [清单文件](../develop/add-in-manifests.md)，它可用于发布加载项。必须单独发布项目的 Web 应用程序文件。本文介绍如何使用 Visual Studio 2015 部署 Web 项目并打包加载项。

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a>使用 Visual Studio 2017 部署 Web 项目

完成以下步骤以使用 Visual Studio 2017 部署 Web 项目。

1. 在**解决方案资源管理器**中，打开加载项项目的快捷菜单，然后选择**发布**。
    
    将显示**发布加载项**页。
    
2. 选择**当前配置文件**下拉列表中的配置文件，或选择**新建…** 以创建新的配置文件。
    
    > [!NOTE]
    > 发布配置文件指定要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。

    如果你选择**新建...**，将会显示**创建发布配置文件**向导。可以使用此向导从托管提供程序（如 Microsoft Azure）的网站导入发布配置文件，或创建新配置文件并添加你的服务器、凭据以及下一过程中的其他设置。
    
    有关导入发布配置文件或创建新发布配置文件的详细信息，请参阅[创建发布配置文件](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)。
    
3. 在**发布加载项**页中，选择**部署 Web 项目**链接。
    
    出现**发布**对话框。有关使用此向导的详细信息，请参阅[如何：在 Visual Studio 中使用“一键式发布”部署 Web 项目](https://msdn.microsoft.com/library/dd465337.aspx)。
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a>使用 Visual Studio 2017 打包加载项的具体步骤

完成以下步骤以使用 Visual Studio 2017 打包加载项。

1. 在**发布加载项**页上，选择**打包加载项**按钮。
    
    **加载项包** 页上将显示向导。
    
2. 在**你的网站托管在何处?** 框中，输入托管加载项内容文件的网站 URL，然后选择**完成**。
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure 网站自动提供 HTTPS 端点。

    此时，Visual Studio 生成发布加载项所需的文件，并打开发布输出文件夹。
    
如果计划将加载项提交到 AppSource，可以选择**执行验证检查**按钮，以发现将会导致加载项被拒绝的任何问题。 应先解决所有问题，再将加载项提交到应用商店。

现在，可以将 XML 清单上传到适当位置，以[发布加载项](../publish/publish.md)。XML 清单位于 `app.publish` 文件夹的 `OfficeAppManifests` 中。例如：

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>另请参阅

- [发布 Office 加载项](../publish/publish.md)
- [将解决方案提交到 AppSource 和 Office 应用商店](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
