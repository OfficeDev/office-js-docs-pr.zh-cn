---
title: Excel 自定义函数教程
description: 在本教程中，你将创建一个 Excel 外接程序，其中包含可执行计算、请求 Web 数据或流式传输 Web 数据的自定义函数。
ms.date: 06/10/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 9550986edcbbed56c69e25e183c304ebe6f6cc07
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091057"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>教程：在 Excel 中创建自定义函数

用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。 Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。 可以创建自定义函数，以执行简单的任务（如计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。

在本教程中，你将：
> [!div class="checklist"]
> - 使用[适用于 Office 加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md)来创建自定义函数加载项。
> - 使用预生成的自定义函数来执行简单计算。
> - 创建从 Web 获取数据的自定义函数。
> - 创建从 Web 传送实时数据的自定义函数。

## <a name="prerequisites"></a>先决条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- 已连接到 Microsoft 365 订阅的 Office (包括 Office 网页版)。

  > [!NOTE]
  > 如果你还没有 Office，可以[加入 Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)以免费获得为期 90 天的可续订 Microsoft 365 订阅，以便在开发期间使用。

## <a name="create-a-custom-functions-project"></a>创建自定义函数项目

 若要开始，请创建代码项目以生成自定义函数加载项。[Office 加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md) 将使用一些可以试用的预生成自定义函数来设置项目。如果已运行自定义函数快速启动并生成了项目，请继续使用该项目，然后改为跳到 [此步骤](#create-a-custom-function-that-requests-data-from-the-web)。

> [!NOTE]
> 如果重新创建 yo office 项目，则可能遇到错误，因为 Office 缓存已具有同名函数的实例。在运行 `npm run start` 之前，可以通过 [清除 Office 缓存](../testing/clear-cache.md) 来防止这种情况。

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **选择项目类型:** `Excel Custom Functions Add-in project`
    - **选择脚本类型:** `JavaScript`
    - **要如何命名加载项?** `My custom functions add-in`

    :::image type="content" source="../images/yo-office-excel-cf-quickstart.png" alt-text="Yeoman Office 加载项生成器命令行界面提示自定义函数项目的屏幕截图。":::

    Yeoman 生成器将创建项目文件并安装支持的 Node 组件。

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. 导航到项目的根文件夹。

    ```command&nbsp;line
    cd "My custom functions add-in"
    ```

1. 生成项目。

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。如果系统在运行 `npm run build` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

1. 启动在 Node.js 中运行的本地 Web 服务器。 可以在 Excel 中试用自定义函数加载项。

# <a name="excel-on-windows-or-mac"></a>[Windows 版或 Mac 版 Excel](#tab/excel-windows)

若要在 Windows 版或 Mac 版 Excel 中测试加载项，请运行以下命令。 运行此命令时，本地 Web 服务器将启动，Excel 将打开并载入加载项。

```command&nbsp;line
npm run start:desktop
```

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

# <a name="excel-on-the-web"></a>[Excel 网页版](#tab/excel-online)

若要在Excel 网页版中测试加载项，请运行以下命令。 运行此命令时，本地 Web 服务器将启动。 将 "{url}" 替换为你拥有权限的 OneDrive 或 SharePoint 库上 Excel 文档的 URL。

[!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

---

## <a name="try-out-a-prebuilt-custom-function"></a>尝试使用预生成的自定义函数

创建的自定义函数项目中包含一些预生成的自定义函数，这些函数在 **./src/functions/functions.js** 文件中定义。 **./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。 你将使用 CONTOSO 命名空间来访问 Excel 中的自定义函数。

接下来，通过完成以下步骤，尝试使用 `ADD` 自定义函数。

1. 在 Excel 中，转至任意单元格并输入 `=CONTOSO`。 请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。

1. 通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。

`ADD` 自定义函数将计算你提供的两个数字的总和，并返回结果 **210**。

[!include[Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>创建从 Web 请求数据的自定义函数

集成来自 Web 的数据是通过自定义函数来扩展 Excel 的好方法。 接下来，需要创建一个名为“`getStarCount`”的自定义函数，显示给定 Github 存储库所拥有的星星数量。

1. 在“**我的自定义函数加载项**”项目中，找到文件 **./src/functions/functions.js**，并在代码编辑器中将其打开。

1. 在 **function.js** 中，添加以下代码。

    ```JS
    /**
      * Gets the star count for a given Github repository.
      * @customfunction 
      * @param {string} userName string name of Github user or organization.
      * @param {string} repoName string name of the Github repository.
      * @return {number} number of stars given to a Github repository.
      */
      async function getStarCount(userName, repoName) {
        try {
          //You can change this URL to any web request you want to work with.
          const url = "https://api.github.com/repos/" + userName + "/" + repoName;
          const response = await fetch(url);
          //Expect that status code is in 200-299 range
          if (!response.ok) {
            throw new Error(response.statusText)
          }
            const jsonResponse = await response.json();
            return jsonResponse.watchers_count;
        }
        catch (error) {
          return error;
        }
      }
    ```

1. 运行以下命令以重新生成项目。

    ```command&nbsp;line
    npm run build
    ```

1. 完成以下步骤（针对 Excel 网页版或者 Windows 版或 Mac 版 Excel），以在 Excel 中重新注册加载项。 必须完成这些步骤，才能使用新函数。

### <a name="excel-on-windows-or-mac"></a>[Windows 版或 Mac 版 Excel](#tab/excel-windows)

1. 关闭 Excel，然后重新打开 Excel。

1. 在 Excel 中，选择“**插入**”选项卡，然后选择“**我的外接程序**”右侧的向下箭头。

    :::image type="content" source="../images/select-insert.png" alt-text="Windows 上 Excel 中插入功能区的屏幕截图，其中突出显示了“我的外接程序”向下箭头。":::

1. 在可用加载项列表中，查找“**开发人员加载项**”部分，并选择“**我的自定义函数加载项**”以进行注册。

    :::image type="content" source="../images/excel-cf-tutorial-register.png" alt-text="Windows 版 Excel 中的“插入”功能区屏幕截图，在“我的加载项”列表中突出显示了 Excel 自定义函数加载项。":::

# <a name="excel-on-the-web"></a>[Excel 网页版](#tab/excel-online)

1. 在 Excel 中，选择“**插入**”选项卡，然后选择“**加载项**”。

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="Excel 网页版中插入功能区的屏幕截图，突出显示“我的加载项”按钮。":::

1. 选择“管理我的加载项”，然后选择“上载我的加载项”。

1. 选择“浏览...”，并导航到 Yeoman 生成器创建的项目的根目录。

1. 依次选择文件“manifest.xml”，“打开”，然后选择“上载”。

1. 尝试使用新函数。 在单元格 **B1** 中，键入文本 **=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")**，然后按 Enter。 你会看到，单元格 **B1** 中的结果便是 [Excel-Custom-Functions Github 存储库](https://github.com/OfficeDev/Excel-Custom-Functions)所获得的星星的当前数目。

---

## <a name="create-a-streaming-asynchronous-custom-function"></a>创建流式处理异步自定义函数

`getStarCount` 函数返回存储库在特定时刻所拥有的星星数量。 自定义函数也会返回不断变化的数据。 这些函数称为流式处理函数。 它们必须包含一个 `invocation` 参数，该参数引用调用该函数的单元格。 `invocation` 参数用于随时更新该单元格的内容。  

在下面的代码示例中，请注意，有两个函数：`currentTime` 和 `clock`。 `currentTime` 函数是不使用流式处理的静态函数。 它将以字符串形式返回日期。 `clock` 函数使用 `currentTime` 函数每秒向 Excel 中的单元格提供一次新时间。 它使用 `invocation.setResult` 将时间传递给 Excel 单元格，并使用 `invocation.onCanceled` 处理函数取消。 

**我的自定义函数加载项** 项目已在 **./src/functions/functions.js** 文件中包含以下两个函数。

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}
    
/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);
    
  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

若要试用函数，请在单元格 **C1** 中键入文本 **=CONTOSO.CLOCK()**，然后按 Enter。 此时会显示当前日期，该日期每秒更新一次。 虽然此时钟只是一个循环计时器，但利用这一理念，你可以在更复杂的函数上设置计时器，以便执行对实时数据的 Web 请求。

## <a name="next-steps"></a>后续步骤

恭喜！ 你已经创建新的自定义函数项目，试用了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了流式传输数据的自定义函数。 接下来，你可以将项目修改为使用共享运行时，使函数更容易与任务窗格交互。 请按照以下文章中的步骤进行操作。

> [!div class="nextstepaction"]
> [配置加载项以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
