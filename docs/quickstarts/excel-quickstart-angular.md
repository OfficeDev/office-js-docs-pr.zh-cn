---
title: 使用 Angular 生成 Excel 任务窗格加载项
description: 了解如何使用 Office JS API 和 Angular 生成简单的 Excel 任务窗格加载项。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 5898d9bd3072e829c35afac90348cb844f96011c
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132317"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a>使用 Angular 生成 Excel 任务窗格加载项

本文将逐步介绍如何使用 Angular 和 Excel JavaScript API 生成 Excel 任务加载项。

## <a name="prerequisites"></a>先决条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>创建加载项项目

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **选择项目类型:** `Office Add-in Task Pane project using Angular framework`
- **选择脚本类型:** `TypeScript`
- **要如何命名加载项?** `My Office Add-in`
- **要支持哪一个 Office 客户端应用程序?** `Excel`

![项目类型设置为“Angular 框架” 的 Yeoman Office 外接程序生成器命令行界面屏幕截图](../images/yo-office-excel-angular-2.png)

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>浏览项目

使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。 如果想要浏览加载项项目的主要组件，请在代码编辑器中打开项目并检查下面列出的文件。 准备好试用加载项时，请转至下一部分。

- 项目根目录中的 **manifest.xml** 文件定义加载项的设置和功能。
- **./src/taskpane/app/app.component.html** 文件包含组成任务窗格的 HTML。
- **./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。
- **./src/taskpane/app/app.component.ts** 文件包含用于加快任务窗格与 Excel 之间的交互的 Office JavaScript API 代码。

## <a name="try-it-out"></a>试用

1. 导航到项目的根文件夹。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. 在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。

    ![Excel 主页菜单的屏幕截图，突出显示“显示任务窗格”按钮](../images/excel-quickstart-addin-3b.png)

4. 选择工作表中的任何一系列单元格。

5. 在任务窗格的底部，选择“**运行**”链接，价格选定范围的颜色设为黄色。

    ![Excel 的屏幕截图，其中“加载项”任务窗格处于打开状态，并且“加载项”任务窗格中突出显示“运行”按钮](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>后续步骤

祝贺，你已使用 Angular 成功创建了 Excel 任务窗格加载项！ 接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。

> [!div class="nextstepaction"]
> [Excel 加载项教程](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>另请参阅

* [Office 加载项平台概述](../overview/office-add-ins.md)
* [开发 Office 加载项](../develop/develop-overview.md)
* [Excel 加载项中的 Word JavaScript 对象模型](../excel/excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
