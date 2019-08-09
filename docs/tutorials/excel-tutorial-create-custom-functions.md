---
title: Excel 自定义函数教程
description: 在本教程中，你将创建一个 Excel 外接程序，其中包含可执行计算、请求 Web 数据或流 Web 数据的自定义函数。
ms.date: 07/09/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 4b74463eafd5ac1b70e59cef6ef1f9f33cf0ffa2
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268178"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>教程：在 Excel 中创建自定义函数

用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。 Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。 可以创建自定义函数，以执行简单的任务（如计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。

在本教程中，你将：
> [!div class="checklist"]
> * 使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建自定义函数加载项。 
> * 使用预生成的自定义函数来执行简单计算。
> * 创建从 Web 获取数据的自定义函数。
> * 创建从 Web 传送实时数据的自定义函数。

## <a name="prerequisites"></a>先决条件

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* Windows (版本1904或更高版本, 连接到 Office 365 订阅) 或 web 上的 Excel

## <a name="create-a-custom-functions-project"></a>创建自定义函数项目

 首先，创建代码项目以构建自定义函数加载项。 [Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)将使用一些预生成的自定义函数来设置您的项目, 您可以试用这些函数。如果已运行自定义函数 "快速启动" 并生成了一个项目, 请继续使用该项目, 然后跳到[此步骤](#create-a-custom-function-that-requests-data-from-the-web)。

1. 运行下面的命令，再回答如下所示的提示问题。
    
    ```command&nbsp;line
    yo office
    ```
    
    * **选择项目类型:** `Excel Custom Functions Add-in project`
    * **选择脚本类型:** `JavaScript`
    * **要如何命名加载项?** `starcount`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/starcountPrompt.png)
    
    Yeoman 生成器将创建项目文件并安装支持的 Node 组件。

2. 导航到项目的根文件夹。
    
    ```command&nbsp;line
    cd starcount
    ```

3. 生成项目。
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行 `npm run build` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

4. 启动在 Node.js 中运行的本地 Web 服务器。 您可以在 Excel 网页或 Windows 中试用自定义函数加载项。

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[Windows 或 Mac 上的 Excel](#tab/excel-windows)

若要在 Excel 中的 Windows 或 Mac 上测试外接程序, 请运行以下命令。 运行此命令时, 本地 web 服务器将启动, 并且 Excel 将在加载的外接程序中打开。

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[在 web 上的 Excel](#tab/excel-online)

若要在 Excel 中的浏览器上测试外接程序, 请运行以下命令。 运行此命令时，本地 Web 服务器将启动。

```command&nbsp;line
npm run start:web
```

若要使用自定义函数外接程序, 请在 web 上的 Excel 中打开一个新工作簿。 在此工作簿中, 完成以下步骤以旁加载您的外接程序。

1. 在 Excel 中, 选择 "**插入**" 选项卡, 然后选择 "**外接程序**"。

   ![在 Excel 中的 "我的外接程序" 图标突出显示的网页中插入功能区](../images/excel-cf-online-register-add-in-1.png)
   
2. 选择“管理我的加载项”****，然后选择“上载我的加载项”****。

3. 选择“浏览...”****，并导航到 Yeoman 生成器创建的项目的根目录。

4. 依次选择文件“manifest.xml”****，“打开”****，然后选择“上载”****。

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>尝试预生成的自定义函数

您创建的自定义函数项目包含一些预生成的自定义函数, 这些函数是在 **/src/functions/functions.js**文件中定义的。 **./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。 你将使用 CONTOSO 命名空间来访问 Excel 中的自定义函数。

接下来，通过完成以下步骤来尝试使用 `ADD` 自定义函数：

1. 在 Excel 中，转至任意单元格并输入 `=CONTOSO`。 请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。

2. 通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。

`ADD` 自定义函数将计算你提供的两个数字的总和，并返回结果 **210**。

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>创建从 Web 请求数据的自定义函数

集成来自 Web 的数据是通过自定义函数来扩展 Excel 的好方法。 接下来, 将创建一个名为`getStarCount`的自定义函数, 该函数显示给定 Github 存储库拥有的星星数。

1. 在**starcount**项目中, 找到 **/src/functions/functions.js**并在代码编辑器中打开该文件。 

2. 在**函数 .js**中, 添加以下代码: 

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

3. 运行以下命令以重建项目。

    ```command&nbsp;line
    npm run build
    ```

4. 完成以下步骤 (适用于 web、Windows 或 Mac 上的 Excel), 以便在 Excel 中重新注册加载项。 您必须完成这些步骤, 新函数才可用。

### <a name="excel-on-windows-or-mactabexcel-windows"></a>[Windows 或 Mac 上的 Excel](#tab/excel-windows)

1. 关闭 Excel，然后重新打开 Excel。

2. 在 Excel 中, 选择 "**插入**" 选项卡, 然后选择位于 **"我的外接程序**" 右侧的向下箭头。 ![在 Excel 中的 "我的外接程序" 箭头突出显示 Windows 中插入功能区](../images/select-insert.png)

3. 在可用加载项列表中, 找到 "**开发人员外**接程序" 部分, 然后选择 " **starcount** " 外接程序进行注册。
    ![在 Windows Excel 中插入带有 "我的外接程序" 列表中突出显示 Excel 自定义函数外接程序的功能区](../images/list-starcount.png)


# <a name="excel-on-the-webtabexcel-online"></a>[在 web 上的 Excel](#tab/excel-online)

1. 在 Excel 中, 选择 "**插入**" 选项卡, 然后选择 "**外接程序**"。 ![在 Excel 中的 "我的外接程序" 图标突出显示的网页中插入功能区](../images/excel-cf-online-register-add-in-1.png)

2. 选择“管理我的加载项”****，然后选择“上载我的加载项”****。

3. 选择“浏览...”****，并导航到 Yeoman 生成器创建的项目的根目录。

4. 依次选择文件“manifest.xml”****，“打开”****，然后选择“上载”****。

---

<ol start="5">
<li> 尝试使用新函数。 在单元格<strong>B1</strong>中, 键入文本<strong>= CONTOSO。GETSTARCOUNT ("OfficeDev", "Excel-自定义函数")</strong> , 然后按 enter。 您应该会看到单元格<strong>B1</strong>中的结果是为[Excel 自定义函数 Github 存储库](https://github.com/OfficeDev/Excel-Custom-Functions)提供的当前星数。</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>创建流式处理异步自定义函数

`getStarCount`函数返回存储库在特定时间点的星数。 自定义函数还可以返回不断变化的数据。 这些函数称为流式处理函数。 它们必须包含一个`invocation`参数, 该参数引用函数的调用位置的单元格。 此`invocation`参数用于在任何时间更新单元格的内容。  

在下面的代码示例中, 您会注意到有两个函数`currentTime`和`clock`。 `currentTime`函数是不使用流式处理的静态函数。 它以字符串的形式返回日期。 `clock`函数使用`currentTime`函数在 Excel 中每秒向单元格提供新时间。 它使用`invocation.setResult`将时间传递到 Excel 单元格, 并`invocation.onCanceled`处理取消该函数时发生的情况。

1. 在 " **starcount** " 项目中, 将以下代码添加到 **./src/functions/functions.js**并保存文件。

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

2. 运行以下命令以重建项目。

    ```command&nbsp;line
    npm run build
    ```

3. 完成以下步骤 (适用于 web、Windows 或 Mac 上的 Excel), 以便在 Excel 中重新注册加载项。 您必须完成这些步骤, 新函数才可用。 

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[Windows 或 Mac 上的 Excel](#tab/excel-windows)

1. 关闭 Excel，然后重新打开 Excel。

2. 在 Excel 中, 选择 "**插入**" 选项卡, 然后选择位于 **"我的外接程序**" 右侧的向下箭头。 ![在 Excel 中的 "我的外接程序" 箭头突出显示 Windows 中插入功能区](../images/select-insert.png)

3. 在可用加载项列表中, 找到 "**开发人员外**接程序" 部分, 然后选择 " **starcount** " 外接程序进行注册。
    ![在 Windows Excel 中插入带有 "我的外接程序" 列表中突出显示 Excel 自定义函数外接程序的功能区](../images/list-starcount.png)

# <a name="excel-on-the-webtabexcel-online"></a>[在 web 上的 Excel](#tab/excel-online)

1. 在 Excel 中, 选择 "**插入**" 选项卡, 然后选择 "**外接程序**"。 ![在 Excel 中的 "我的外接程序" 图标突出显示的网页中插入功能区](../images/excel-cf-online-register-add-in-1.png)

2. 选择“管理我的加载项”****，然后选择“上载我的加载项”****。

3. 选择“浏览...”****，并导航到 Yeoman 生成器创建的项目的根目录。

4. 依次选择文件“manifest.xml”****，“打开”****，然后选择“上载”****。

--- 

<ol start="4">
<li>尝试使用新函数。 在单元格<strong>C1</strong>中, 键入文本<strong>= CONTOSO。CLOCK ())</strong> , 然后按 enter。 您应看到当前日期, 该日期每秒处理一次更新。 虽然此时钟只是循环中的计时器, 但在对实时数据发出 web 请求的更复杂的函数上设置计时器时, 可以使用相同的想法。</li>
</ol>

## <a name="next-steps"></a>后续步骤

恭喜！ 您已创建新的自定义函数项目, 并试用一个预建函数, 创建了一个从 web 请求数据的自定义函数, 并创建了一个流式处理数据的自定义函数。 您也可以尝试使用[自定义函数调试指令](../excel/custom-functions-debugging.md)来调试此函数。 若要详细了解 Excel 中的自定义函数，请继续阅读以下文章：

> [!div class="nextstepaction"]
> [在 Excel 中创建自定义函数](../excel/custom-functions-overview.md)
