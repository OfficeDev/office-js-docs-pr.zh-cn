---
title: 使用 Yeoman 生成器创建 Office 加载项项目
description: 了解如何使用 Office 外接程序的 Yeoman 生成器创建 Office 外接程序项目。
ms.date: 08/19/2022
ms.localizationpriority: high
ms.openlocfilehash: f109c4dbc386a4cc23f72d0c67f9e4904360bba4
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422785"
---
# <a name="create-office-add-in-projects-using-the-yeoman-generator"></a>使用 Yeoman 生成器创建 Office 加载项项目

Office 外接程序的 [Yeoman 生成器](https://github.com/OfficeDev/generator-office) (也称为“Yo Office”) 是一种基于Node.js的交互式命令行工具，用于创建 Office 外接程序开发项目。 我们建议使用此工具创建外接程序项目，除非希望外接程序的服务器端代码位于其中。基于 NET 的语言 (（例如 C# 或 VB.Net) ）或者希望在 Internet Information Server (IIS) 中托管加载项。 在后两种情况下， [使用 Visual Studio 创建加载项](develop-add-ins-visual-studio.md)。

工具创建的项目具有以下特征。

- 他们有一个包含 **package.json** 文件的标准 [npm](https://www.npmjs.com/) 配置。
- 其中包括几个有用的脚本，用于生成项目、启动服务器、在 Office 中旁加载加载项以及其他任务。
- 他们使用 [Webpack](https://webpack.js.org/) 作为捆绑程序和基本任务运行程序。
- 在开发模式下，它们由基于 webpack 的基于 Node.js 的 webpack-dev-server 托管在 localhost 上，这是一种面向开发的 [快速](http://expressjs.com/) 服务器版本，支持热重载和在更改时重新编译。
- 默认情况下，工具会安装所有依赖项，但可以使用命令行参数推迟安装。
- 其中包括完整的加载项清单。
- 他们有一个“Hello World”级加载项，该加载项在工具完成后即可运行。
- 它们包括一个多填充和一个配置为将 TypeScript 和最新版本的 JavaScript 转译到 ES5 JavaScript 的转译器。 这些功能可确保 Office 加载项可能运行的所有运行时（包括 Internet Explorer）都支持外接程序。

> [!TIP]
> 如果想要明显偏离这些选择，例如使用不同的任务运行程序或不同的服务器，建议在运行该工具时选择 [“仅限清单”选项](#manifest-only-option)。

## <a name="install-the-generator"></a>安装生成器

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="use-the-tool"></a>使用该工具

在系统提示符中使用以下命令启动该工具， (不是 bash 窗口) 。

```command&nbsp;line
yo office 
```

需要加载很多，因此可能需要 20 秒才能启动该工具。 该工具会向你提出一系列问题。 对于某些用户，只需键入提示的答案。 对于其他人，你将获得可能的答案列表。 如果给定列表，请选择一个，然后选择 Enter。

第一个问题要求你在六种类型的项目之间进行选择。 选项包括：

- **Office 加载项任务窗格项目**
- **使用Angular框架的 Office 加载项任务窗格项目**
- **使用React框架的 Office 加载项任务窗格项目**
- **支持单一登录的 Office 加载项任务窗格项目**
- **仅包含清单的 Office 外接程序项目**
- **Excel 自定义函数外接程序项目**

![显示 Yeoman 生成器中项目类型的提示和可能的答案的屏幕截图。](../images/yo-office-project-type-prompt.png)

> [!NOTE]
> **支持单一登录选项的 Office 加载项任务窗格项目** 生成一个项目，可用于查看单一登录 (SSO) 在加载项中的工作原理。 项目不能用作生产外接程序的基础。 若要获取可作为生产加载项基础的已启用 SSO 的项目，请参阅 [示例存储库中某个 SSO 示例的](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth)“完整”版本。
>
> **包含仅限清单选项的 Office 外接程序项目** 生成包含基本加载项清单和最小基架的项目。 有关该选项的详细信息，请参阅 [仅限清单的选项](#manifest-only-option)。

下一个问题要求在 **TypeScript** 和 **JavaScript** 之间进行选择。  (如果在前面的问题中选择了“仅限清单”选项，则跳过此问题。) 

![屏幕截图显示用户在上述问题中选择了“Office 加载项任务窗格项目”，并在 Yeoman 生成器中显示语言提示以及可能的答案 TypeScript 和 JavaScript。](../images/yo-office-language-prompt.png)

然后，系统会提示你为外接程序指定一个名称。 指定的名称将用于加载项的清单中，但稍后可以对其进行更改。

![显示用户选择 TypeScript 作为上一个问题的屏幕截图，并显示 Yeoman 生成器中加载项名称的提示。](../images/yo-office-name-prompt.png)

然后，系统会提示你选择外接程序应在哪个 Office 应用程序中运行。 有六个可能的应用程序可供选择： **Excel**、 **OneNote**、 **Outlook**、 **PowerPoint**、 **Project** 和 **Word**。 必须只选择一个，但以后可以更改清单以支持其他 Office 应用程序。 例外情况是 Outlook。 支持 Outlook 的清单不能支持任何其他 Office 应用程序。

![屏幕截图显示用户将项目命名为“Contoso 外接程序”，并在 Yeoman 生成器中显示 Office 应用程序提示和可能的答案。](../images/yo-office-host-prompt.png)

回答此问题后，生成器将创建项目并安装依赖项。 你可能会在屏幕上的 npm 输出中看到 **WARN** 消息。 可以忽略这些。 还可能会看到发现漏洞的消息。 可以暂时忽略这些内容，但最终需要在外接程序发布到生产环境之前对其进行修复。 有关修复漏洞的详细信息，请打开浏览器并搜索“npm 漏洞”。

如果创建成功，你将看到 **一个恭喜！** 命令窗口中的消息，后跟一些建议的后续步骤。  (如果将生成器用作快速入门或教程的一部分，请忽略命令窗口中的后续步骤并继续执行文章中的说明。) 

> [!TIP]
> 如果要创建 Office 外接程序项目的基架，但推迟安装依赖项，请将选项 `--skip-install` 添加到 `yo office` 命令。 以下代码是一个示例。
>
> ```command&nbsp;line
> yo office --skip-install
> ```
>
> 准备好安装依赖项后，在命令提示符下导航到项目的根文件夹并输入 `npm install`。

## <a name="manifest-only-option"></a>仅限清单选项

此选项仅为加载项创建清单。 生成的项目没有Hello World加载项、任何脚本或任何依赖项。 在以下方案中使用此选项。

- 你希望使用与 Yeoman 生成器项目默认安装和配置的工具不同的工具。 例如，需要使用不同的捆绑程序、转译器、任务运行程序或开发服务器。
- 你希望使用 Web 应用程序开发框架，而不是Angular或React，例如 Vue。

有关将生成器与仅限清单选项配合使用的示例，请参阅 [使用 Vue 生成 Excel 任务窗格加载项](../quickstarts/excel-quickstart-vue.md)。

## <a name="use-command-line-parameters"></a>使用命令行参数

还可以向 `yo office` 命令添加参数。 两种最常用的查询为：

- `yo office --details`：这将输出有关所有其他命令行参数的简短帮助。
- `yo office --skip-install`：这将阻止生成器安装依赖项。

有关命令行参数的详细参考，请参阅适用于 Office 外接程序的 [Yeoman 生成器上的生成器的](https://github.com/officedev/generator-office)自述文件。

## <a name="troubleshooting"></a>疑难解答

如果在使用该工具时遇到问题，则第一步应该是重新安装它，以确保你拥有最新版本。  (请参阅 [安装生成器](#install-the-generator) 以获取详细信息。) 如果这样做无法解决问题，请搜索 [该工具的 GitHub 存储库问题](https://github.com/OfficeDev/generator-office/issues) ，查看是否有其他人遇到过相同的问题并找到了解决方案。 如果没有人，请 [创建新问题](https://github.com/OfficeDev/generator-office/issues/new?assignees=&labels=needs+triage&template=bug_report.md&title=)。
