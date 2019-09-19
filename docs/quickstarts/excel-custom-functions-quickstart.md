---
ms.date: 09/18/2019
description: 在 Excel 快速入门指南中开发自定义函数。
title: 自定义功能快速入门
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f34a8817a7c8ef2679fc8ce0a6ad17cec600531b
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035327"
---
# <a name="get-started-developing-excel-custom-functions"></a>开始开发 Excel 自定义函数

通过自定义函数，开发人员现在可以通过在 JavaScript 或 Typescript 中将新函数定义为外接程序的一部分，将它们添加到 Excel 中。 Excel 用户可以像对待 Excel 中的任何本机函数一样访问自定义函数，例如`SUM()`。

## <a name="prerequisites"></a>先决条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Windows 上的 Excel （版本1904或更高版本，连接到 Office 365 订阅）或 web 上的 Excel
* Office on Mac （连接到 Office 365 订阅）支持 Excel 自定义函数，本教程的更新即将推出。

>[!NOTE]
>Excel 自定义函数在 Office 2019 中不受支持（一次性购买）。

## <a name="build-your-first-custom-functions-project"></a>生成第一个自定义函数项目

首先，使用 Yeoman 生成器创建自定义函数项目。 这将为你的项目设置开始对自定义函数进行编码所需的正确文件夹结构、源文件和依存关系。

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **选择项目类型:** `Excel Custom Functions Add-in project`
    - **选择脚本类型:** `JavaScript`
    - **要如何命名加载项?** `starcount`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/starcountPrompt.png)

    Yeoman 生成器将创建项目文件并安装支持的 Node 组件。

2. Yeoman 生成器将为您提供有关如何处理项目的命令行中的一些说明，但忽略它们并继续按照我们的说明操作。 导航到项目的根文件夹。

    ```command&nbsp;line
    cd starcount
    ```

3. 生成项目。 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行 `npm run build` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

4. 启动在 Node.js 中运行的本地 Web 服务器。 您可以在 Excel 网页或 Windows 中试用自定义函数加载项。 系统可能会提示您打开加载项的任务窗格，但这是可选的。 您仍可以运行自定义函数，而无需打开加载项的任务窗格。

# <a name="excel-on-windowstabexcel-windows"></a>[Windows 上的 Excel](#tab/excel-windows)

若要在 Windows 中的 Excel 中测试外接程序，请运行以下命令。 运行此命令时，本地 web 服务器将启动，并且 Excel 将在加载的外接程序中打开。

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[在 web 上的 Excel](#tab/excel-online)

若要在 Excel 网页上测试您的外接程序，请运行以下命令。 运行此命令时，本地 Web 服务器将启动。

```command&nbsp;line
npm run start:web
```

若要使用自定义函数外接程序，请在 Excel 中的浏览器中打开一个新工作簿。 在此工作簿中，完成以下步骤以旁加载您的外接程序。

1. 在 Excel 中，选择 "**插入**" 选项卡，然后选择 "**外接程序**"。

   ![在 Excel 中的 "我的外接程序" 图标突出显示的网页中插入功能区](../images/excel-cf-online-register-add-in-1.png)
   
2. 选择“管理我的加载项”****，然后选择“上载我的加载项”****。

3. 选择“浏览...”****，并导航到 Yeoman 生成器创建的项目的根目录。

4. 依次选择文件“manifest.xml”****，“打开”****，然后选择“上载”****。

---

## <a name="try-out-a-prebuilt-custom-function"></a>尝试预生成的自定义函数

使用 Yeoman 生成器创建的自定义函数项目包含一些预生成的自定义函数，这些函数是在 **/src/functions/functions.js**文件中定义的。 项目根目录中的 **/manifest.xml**文件指定所有自定义函数均属于该`CONTOSO`命名空间。

在 Excel 工作簿中，通过完成`ADD`以下步骤来尝试使用自定义函数：

1. 选择一个单元格并`=CONTOSO`键入。 请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。

2. 通过在`CONTOSO.ADD`单元格中键入值`10` `=CONTOSO.ADD(10,200)`并`200`按 enter 来运行函数，并使用数字和作为输入参数。

`ADD` 自定义函数计算指定为输入参数的两个数字的总和。 键入 `=CONTOSO.ADD(10,200)` 应在按下 Enter 后在单元格中生成结果 **210**。

## <a name="next-steps"></a>后续步骤

恭喜！你已成功在 Excel 加载项中创建了自定义函数！ 接下来，使用流式数据功能生成更复杂的加载项。 下面的链接将指导您完成 Excel 加载项的自定义函数教程中的后续步骤。

> [!div class="nextstepaction"]
> [Excel 自定义函数加载项教程](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a>另请参阅

* [自定义函数概述](../excel/custom-functions-overview.md)
* [自定义函数元数据](../excel/custom-functions-json.md)
* [Excel 自定义函数的运行时](../excel/custom-functions-runtime.md)