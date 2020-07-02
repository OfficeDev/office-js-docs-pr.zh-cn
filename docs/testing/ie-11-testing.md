---
ms.date: 05/16/2020
description: 使用 Internet Explorer 11 测试 Office 外接程序。
title: Internet Explorer 11 测试
localization_priority: Normal
ms.openlocfilehash: 1d6852d08308088a020e86ce7f5ab9cfdb9ab978
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006435"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a>使用 Internet Explorer 11 测试 Office 外接程序

根据你的外接程序的规范，你可能会计划支持较早版本的 Windows 和 Office，这需要在 Internet Explorer 11 上进行测试。 在将外接程序提交到 AppSource 时，通常需要执行此过程。 您可以使用以下命令行工具从外接程序使用的更新式运行时切换到 Internet Explorer 11 运行时进行此测试。

## <a name="pre-requisites"></a>先决条件

- [Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）
- 一个代码编辑器。 建议[Visual Studio Code](https://code.visualstudio.com/)
- [是 Office 预览体验计划的一部分](https://insider.office.com)

这些说明假定您先设置了 "Yo Office 生成器" 项目。 如果你之前未执行此操作，请考虑阅读快速启动，例如， [Excel 外接程序](../quickstarts/excel-quickstart-jquery.md)。

## <a name="using-ie11-tooling"></a>使用 IE11 工具

1. 创建 "Yo Office 生成器" 项目。 无论选择哪种类型的项目，此工具都将适用于所有项目类型。

> !便笺如果您有一个现有项目，并且想要添加此工具而不创建新项目，请跳过此步骤并移动到下一步。 

2. 在新项目的根文件夹中，在命令行中运行以下命令：

```command&nbsp;line
npx office-addin-dev-settings webview manifest.xml ie
```
您应该会在命令行中看到一条注释，web 视图类型现在设置为 IE。

> !尖不必使用此工具，但应帮助调试与 Internet Explorer 11 运行时相关的大多数问题。 为实现全面的可靠性，应使用安装了 Windows 7 和 Office 2013 副本的计算机进行测试。

## <a name="command-settings"></a>命令设置

如果您有一个不同的清单路径，请在命令中指定此路径，如下所示：

`npx office-addin-dev-settings webview [path to your manifest] ie`

该 `office-addin-dev-settings webview` 命令还可以采用若干个运行时作为参数：

- 限于
- 距
-  默认值

## <a name="see-also"></a>另请参阅
* [测试和调试 Office 加载项](test-debug-office-add-ins.md)
* [旁加载 Office 外接程序进行测试](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [使用 Windows 10 上的开发人员工具调试加载项](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [从任务窗格附加调试器](attach-debugger-from-task-pane.md)
