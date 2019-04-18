---
ms.date: 03/06/2019
description: 在 Excel 快速入门指南中开发自定义函数。
title: 自定义函数快速入门 (预览)
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 80c500e1e30e8751a7d969d33cd7e13b7943b1b5
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914296"
---
# <a name="get-started-developing-excel-custom-functions"></a>开始开发 Excel 自定义函数

通过自定义函数, 开发人员现在可以通过在 JavaScript 或 Typescript 中将新函数定义为外接程序的一部分, 将它们添加到 Excel 中。 excel 用户可以像对待 excel 中的任何本机函数一样访问自定义函数, 例如`SUM()`。

## <a name="prerequisites"></a>先决条件

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

您需要以下工具和相关资源来开始创建自定义函数。

- [Node.js](https://nodejs.org/en/)（版本 8.0.0 或更高版本）

- [Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）

- 最新版本的 [Yeoman](https://yeoman.io/) 和[适用于 Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > 即使以前安装了 Yeoman 生成器, 我们也建议您将程序包从 npm 更新到最新版本。

## <a name="build-your-first-custom-functions-project"></a>生成第一个自定义函数项目

首先，使用 Yeoman 生成器创建自定义函数项目。 这将为你的项目设置开始对自定义函数进行编码所需的正确文件夹结构、源文件和依存关系。

1. 运行下面的命令，再回答如下所示的提示问题。

    ```
    yo office
    ```

    - 选择项目类型：`Excel Custom Functions Add-in project (...)`

    - 选择脚本类型：`JavaScript`

    - 要如何命名加载项？ `stock-ticker`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/12-10-fork-cf-pic.jpg)

    Yeoman 生成器将创建项目文件并安装支持的 Node 组件。

2. 导航到刚创建的项目文件夹。

    ```
    cd stock-ticker
    ```

3. 信任自签名证书, 您需要运行此项目。 有关适用于 Windows 或 Mac 的详细说明，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。  

4. 生成项目。

    ```
    npm run build
    ```

5. 启动在 Node.js 中运行的本地 Web 服务器。

    - 如果使用 Excel for Windows 测试自定义函数, 请运行以下命令来启动本地 web 服务器, 启动 Excel, 并旁加载外接程序:

        ```
         npm run start
        ```
        运行此命令后, 命令提示符将显示有关启动 web 服务器的详细信息。 Excel 将从加载的加载项开始。 如果加载项未加载，请检查是否已正确完成步骤 3。

    - 如果使用 Excel Online 测试自定义函数, 请运行以下命令来启动本地 web 服务器:

        ```
        npm run start-web
        ```

         运行此命令后, 命令提示符将显示有关启动 web 服务器的详细信息。 若要使用您的函数, 请在 Excel Online 中打开一个新工作簿。 在此工作簿中, 需要加载外接程序。 

        若要执行此操作, 请选择功能区上的 "**插入**" 选项卡, 然后选择 "**获取外接程序**"。在生成的新窗口中, 确保您在 "**我的外接程序**" 选项卡上。接下来, 选择 "**管理我的外接程序" > 上传我的外接程序**。 浏览清单文件并将其上传。 如果加载项未加载, 请检查是否已正确完成步骤3。

## <a name="try-out-the-prebuilt-custom-functions"></a>尝试预生成的自定义函数

使用 Yeoman 生成器创建的自定义函数项目包含一些预生成的自定义函数，这些函数在 **src/customfunction.js** 文件中定义。 项目根目录中的 **manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 名称空间。

在 Excel 工作簿中, 通过完成`ADD`以下步骤来尝试使用自定义函数:

1. 选择一个单元格并`=CONTOSO`键入。 请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。

2. 通过在`CONTOSO.ADD`单元格中键入值`10` `=CONTOSO.ADD(10,200)`并`200`按 enter 来运行函数, 并使用数字和作为输入参数。

`ADD` 自定义函数计算指定为输入参数的两个数字的总和。 键入 `=CONTOSO.ADD(10,200)` 应在按下 Enter 后在单元格中生成结果 **210**。

## <a name="next-steps"></a>后续步骤

恭喜! 你已成功在 Excel 加载项中创建了自定义函数! 接下来, 使用流式数据功能生成更复杂的加载项。 下面的链接将指导您完成 Excel 加载项的自定义函数教程中的后续步骤。

> [!div class="nextstepaction"]
> [Excel 自定义函数加载项教程](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a>另请参阅

* [自定义函数概述](../excel/custom-functions-overview.md)
* [自定义函数元数据](../excel/custom-functions-json.md)
* [Excel 自定义函数的运行时](../excel/custom-functions-runtime.md)
* [自定义函数最佳实践](../excel/custom-functions-best-practices.md)
