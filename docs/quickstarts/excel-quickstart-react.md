---
title: 使用 React 生成 Excel 任务窗格加载项
description: ''
ms.date: 05/02/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 1c0f2f4af1ee14aaf7d581509733e26013657590
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308027"
---
# <a name="build-an-excel-task-pane-add-in-using-react"></a>使用 React 生成 Excel 任务窗格加载项

本文将逐步介绍如何使用 React 和 Excel JavaScript API 生成 Excel 任务加载项。

## <a name="prerequisites"></a>先决条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>创建加载项项目

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

使用 Yeoman 生成器创建 Excel 加载项项目。 运行下面的命令，再回答如下所示的提示问题：

```command&nbsp;line
yo office
```

- **选择项目类型:** `Office Add-in Task Pane project using React framework`
- **选择脚本类型:** `TypeScript`
- **要如何命名加载项?** `my-office-add-in`
- **要支持哪一个 Office 客户端应用程序?** `Excel`

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

## <a name="explore-the-project"></a>浏览项目

使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。 如果想要浏览加载项项目的主要组件，请在代码编辑器中打开项目并检查下面列出的文件。 准备好试用加载项时，请转至下一部分。

- 项目根目录中的 **manifest.xml** 文件定义加载项的设置和功能。
- **./src/taskpane/taskpane.html** 文件定义任务窗格的 HTML 框架，而 **./src/taskpane/components** 文件夹内的文件定义任务窗格 UI 的各个部分。
- **./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。
- **./src/taskpane/components/App.tsx** 文件包含用于加快任务窗格与 Excel 之间的交互的 Office JavaScript API 代码。

## <a name="try-it-out"></a>试用

1. 导航到项目的根文件夹。

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. 在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。

    ![Excel 加载项按钮](../images/excel-quickstart-addin-3b.png)

4. 选择工作表中的任何一系列单元格。

5. 在任务窗格的底部，选择“**运行**”链接，价格选定范围的颜色设为黄色。

    ![Excel 加载项](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>后续步骤

祝贺，你已使用 React 成功创建了 Excel 任务窗格加载项！ 接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。

> [!div class="nextstepaction"]
> [Excel 加载项教程](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>另请参阅

* [Excel 加载项教程](../tutorials/excel-tutorial-create-table.md)
* [Excel JavaScript API 基本编程概念](../excel/excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API 参考](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
