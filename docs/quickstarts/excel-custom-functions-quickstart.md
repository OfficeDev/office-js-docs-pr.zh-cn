---
ms.date: 12/28/2021
description: 在 Excel 中开发自定义函数快速入门指南。
title: 自定义函数快速入门
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 2f4a2ed07c23c3ced19632b9dbfee2957f0f5ba0
ms.sourcegitcommit: b46d2afc92409bfc6612b016b1cdc6976353b19e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/30/2021
ms.locfileid: "61648000"
---
# <a name="get-started-developing-excel-custom-functions"></a>开始开发 Excel 自定义函数

借助自定义函数，开发人员现在可以在 Excel 中添加新函数，方法是在 JavaScript 或 Typescript 中将这些函数定义为加载项的一部分。 Excel 用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。

## <a name="prerequisites"></a>先决条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Windows 版 Excel（版本 1904 或更高版本）或 Excel 网页版。
- Mac 版 Office（关联至 Microsoft 365 订阅）支持 Excel 自定义函数，并且本教程即将推出相关更新。

## <a name="build-your-first-custom-functions-project"></a>生成首个自定义函数项目

首先，使用 Yeoman 生成器创建自定义函数项目。 这将为你的项目设置开始对自定义函数进行编码所需的正确文件夹结构、源文件和依存关系。

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **选择项目类型:** `Excel Custom Functions Add-in project`
    - **选择脚本类型:** `JavaScript`
    - **要如何命名加载项?** `starcount`

    ![Yeoman Office 加载项生成器命令行界面提示自定义函数项目的屏幕截图。](../images/starcountPrompt.png)

    Yeoman 生成器将创建项目文件并安装支持的 Node 组件。

1. Yeoman 生成器将在命令行中为你提供有关如何处理项目的说明，但请忽略它们并继续按照我们的说明进行操作。导航到项目的根文件夹。

    ```command&nbsp;line
    cd starcount
    ```

1. 生成项目。

    ```command&nbsp;line
    npm run build
    ```

1. 启动在 Node.js 中运行的本地 Web 服务器。 你可以在 Excel 网页版或 Windows 版 Excel 中尝试使用自定义函数加载项。 系统可能会提示你打开加载项的任务窗格，不过这是可选的。 你仍可在不打开加载项的任务窗格的情况下运行自定义函数。

# <a name="excel-on-windows"></a>[Windows 版 Excel](#tab/excel-windows)

若要在 Windows 版 Excel 中测试加载项，请运行以下命令。 运行此命令时，本地 Web 服务器将启动，Excel 将打开并载入加载项。

```command&nbsp;line
npm run start:desktop
```

> [!NOTE]
> Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行 `npm run start` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。
    
# <a name="excel-on-the-web"></a>[Excel 网页版](#tab/excel-online)

若要在Excel 网页版中测试加载项，请运行以下命令。 运行此命令时，本地 Web 服务器将启动。

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行 `npm run start` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

若要使用自定义函数加载项，请在 Excel 网页版中打开一个新工作簿。 在此工作簿中，完成以下步骤以旁加载你的加载项。

1. 在 Excel 中，选择“**插入**”选项卡，然后选择“**加载项**”。

   ![Excel 网页版中插入功能区的屏幕截图，突出显示“我的加载项”按钮。](../images/excel-cf-online-register-add-in-1.png)

1. 选择“管理我的加载项”，然后选择“上载我的加载项”。

1. 选择“浏览...”，并导航到 Yeoman 生成器创建的项目的根目录。

1. 依次选择文件“manifest.xml”，“打开”，然后选择“上载”。

---

## <a name="try-out-a-prebuilt-custom-function"></a>尝试使用预生成的自定义函数

使用 Yeoman 生成器创建的自定义函数项目包含一些预生成的自定义函数，这些函数在 **./src/functions/functions.js** 文件中定义。 项目根目录中的 **./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。

在 Excel 工作簿中，通过完成以下步骤尝试使用 `ADD` 自定义函数。

1. 选择单元格并键入 `=CONTOSO`。请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。

1. 通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。

`ADD` 自定义函数计算指定为输入参数的两个数字的总和。 键入 `=CONTOSO.ADD(10,200)` 应在按下 Enter 后在单元格中生成结果 **210**。

## <a name="next-steps"></a>后续步骤

祝贺你，你已成功在 Excel 加载项中创建自定义函数！ 接下来，可生成具有流式数据功能的更复杂的加载项。 通过以下链接，可了解 Excel 自定义函数加载项教程中的后续步骤。

> [!div class="nextstepaction"]
> [Excel 自定义函数加载项教程](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web)

## <a name="troubleshooting"></a>疑难解答

如果多次运行快速入门，可能会遇到问题。 如果 Office 缓存已具有同名函数的实例，则加载项在旁加载时会收到错误。 在运行 `npm run start` 之前，可以通过[清除 Office 缓存](../testing/clear-cache.md)来阻止此操作。

:::image type="content" source="../images/custom-function-already-exists-error.png" alt-text="Excel 中标题为“安装函数时出错”的错误消息。它包含文本“未安装此加载项，因为已存在同名的自定义函数”。":::

## <a name="see-also"></a>另请参阅

- [自定义函数概述](../excel/custom-functions-overview.md)
- [自定义函数元数据](../excel/custom-functions-json.md)
- [Excel 自定义函数的运行时](../excel/custom-functions-runtime.md)
