---
title: Excel 自定义函数教程
description: 在本教程中，你将创建一个 Excel 外接程序，其中包含可执行计算、请求 Web 数据或流 Web 数据的自定义函数。
ms.date: 05/08/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: ed9f16bdb330aa3f092e7d437ccfad6e056e07d4
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952192"
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

* Windows 上的 Excel (64 位版本1810或更高版本) 或 Excel Online

* 加入 [Office 预览体验计划](https://products.office.com/office-insider)（**预览体验成员**级别 - 以前称为“预览体验成员 - 快”）

## <a name="create-a-custom-functions-project"></a>创建自定义函数项目

 首先，创建代码项目以构建自定义函数加载项。 [Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)将使用一些预生成的自定义函数来设置您的项目, 您可以试用这些函数。如果已运行自定义函数 "快速启动" 并生成了一个项目, 请继续使用该项目, 然后跳到[此步骤](#create-a-custom-function-that-requests-data-from-the-web)。

1. 运行下面的命令，再回答如下所示的提示问题。
    
    ```command&nbsp;line
    yo office
    ```
    
    * **选择项目类型:** `Excel Custom Functions Add-in project (...)`
    * **选择脚本类型:** `JavaScript`
    * **要如何命名加载项?** `stock-ticker`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/yo-office-excel-cf.png)
    
    Yeoman 生成器将创建项目文件并安装支持的 Node 组件。

2. 导航到项目的根文件夹。
    
    ```command&nbsp;line
    cd stock-ticker
    ```

3. 生成项目。
    
    ```command&nbsp;line
    npm run build
    ```

4. 启动在 Node.js 中运行的本地 Web 服务器。 可以在 Windows 或 Excel Online 上试用 Excel 中的自定义函数加载项。

# <a name="excel-on-windowstabexcel-windows"></a>[Windows 上的 Excel](#tab/excel-windows)

若要在 Windows 中的 Excel 中测试外接程序, 请运行以下命令。 运行此命令时, 本地 web 服务器将启动, 并且 Windows 上的 Excel 将在加载的外接程序中打开。

```command&nbsp;line
npm run start:desktop
```

> [!NOTE]
> Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行 `npm run start:desktop` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

若要在 Excel Online 中测试外接程序, 请运行以下命令。 运行此命令时，本地 Web 服务器将启动。

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。 如果系统在运行 `npm run start:web` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。

若要使用自定义函数外接程序, 请在 Excel Online 中打开一个新工作簿。 在此工作簿中, 完成以下步骤以旁加载您的外接程序。

1. 在 Excel Online 中，依次选择“插入”**** 选项卡和“加载项”****。

   ![在 Excel Online 中插入带突出显示 "我的外接程序" 图标的功能区](../images/excel-cf-online-register-add-in-1.png)
   
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

集成来自 Web 的数据是通过自定义函数来扩展 Excel 的好方法。 接下来，你将创建一个名为 `stockPrice` 的自定义函数，该函数从 Web API 获取股票报价并将结果返回到工作表的单元格。 你将使用使用 IEX Trading API，该 API 是免费的，并且不需要身份验证。

1. 在**股票报价**项目中, 找到 **/src/functions/functions.js**并在代码编辑器中打开该文件。

2. 在**函数 .Js**中, 找到`increment`函数并在该函数后面添加以下代码。

    ```js
    /**
    * Fetches current stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @returns {number} The current stock price.
    */
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }
    CustomFunctions.associate("STOCKPRICE", stockPrice);
    ```

    `CustomFunctions.associate` 代码会将函数的 `id` 与 JavaScript 中的 `stockPrice` 的函数地址相关联，以便 Excel 能够调用你的函数。

3. 运行以下命令以重建项目。

    ```command&nbsp;line
    npm run build
    ```

4. 完成以下步骤 (针对 Windows 或 Excel Online 上的 Excel), 以便在 Excel 中重新注册加载项。 您必须完成这些步骤, 新函数才可用。 

# <a name="excel-on-windowstabexcel-windows"></a>[Windows 上的 Excel](#tab/excel-windows)

1. 关闭 Excel，然后重新打开 Excel。

2. 在 Excel 中, 选择 "**插入**" 选项卡, 然后选择位于 **"我的外接程序**" 右侧的向下箭头。 ![在 Excel 中的 "我的外接程序" 箭头突出显示 Windows 中插入功能区](../images/select-insert.png)

3. 在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。
    ![在 Windows Excel 中插入带有 "我的外接程序" 列表中突出显示 Excel 自定义函数外接程序的功能区](../images/list-stock-ticker-red.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. 在 Excel Online 中，选择“插入”**** 选项卡，然后选择“加载项”****。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)

2. 选择“管理我的加载项”****，然后选择“上载我的加载项”****。 

3. 选择“浏览...”****，并导航到 Yeoman 生成器创建的项目的根目录。 

4. 依次选择文件“manifest.xml”****，“打开”****，然后选择“上载”****。

---

<ol start="5">
<li> 尝试使用新函数。 在单元格 <strong>B1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICE("MSFT")</strong>，然后按 Enter。 应看到单元格 <strong>B1</strong> 中的结果是 Microsoft 一股股票的当前股票价格。</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>创建流式处理异步自定义函数

`stockPrice` 函数将返回特定时刻的股票价格，但股票价格一直在变化。 接下来，将创建一个名为 `stockPriceStream` 的自定义函数，该函数每隔 1000 毫秒获取一次股票价格。

1. 在**股票报价**项目中, 将以下代码添加到 **。/src/functions/functions.js**并保存文件。

    ```js
    /**
    * Streams real time stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @param {CustomFunctions.StreamingInvocation<number>} invocation
    */
    function stockPriceStream(ticker, invocation) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    invocation.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    invocation.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        invocation.onCanceled = () => {
            clearInterval(timer);
        };
    }
    CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
    ```
    
    `CustomFunctions.associate` 代码会将函数的 `id` 与 JavaScript 中的 `stockPriceStream` 的函数地址相关联，以便 Excel 能够调用你的函数。
    
2. 运行以下命令以重建项目。

    ```command&nbsp;line
    npm run build
    ```

3. 完成以下步骤 (针对 Windows 或 Excel Online 上的 Excel), 以便在 Excel 中重新注册加载项。 您必须完成这些步骤, 新函数才可用。 

# <a name="excel-on-windowstabexcel-windows"></a>[Windows 上的 Excel](#tab/excel-windows)

1. 关闭 Excel，然后重新打开 Excel。

2. 在 Excel 中, 选择 "**插入**" 选项卡, 然后选择位于 **"我的外接程序**" 右侧的向下箭头。 ![在 Excel 中的 "我的外接程序" 箭头突出显示 Windows 中插入功能区](../images/select-insert.png)

3. 在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。
    ![在 Windows Excel 中插入带有 "我的外接程序" 列表中突出显示 Excel 自定义函数外接程序的功能区](../images/list-stock-ticker-red.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. 在 Excel Online 中，选择“插入”**** 选项卡，然后选择“加载项”****。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)

2. 选择“管理我的加载项”****，然后选择“上载我的加载项”****。

3. 选择“浏览...”****，并导航到 Yeoman 生成器创建的项目的根目录。

4. 依次选择文件“manifest.xml”****，“打开”****，然后选择“上载”****。

--- 

<ol start="4">
<li>尝试使用新函数。 在单元格 <strong>C1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong>，然后按 Enter。 假设股票市场开盘，应该会看到单元格 <strong>C1</strong> 中的结果在不断更新，以反映 Microsoft 一股股票的实时价格。</li>
</ol>

## <a name="next-steps"></a>后续步骤

恭喜！ 你已经创建新的自定义函数项目，尝试了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 传送实时数据的自定义函数。 您也可以尝试使用[自定义函数调试指令](../excel/custom-functions-debugging.md)来调试此函数。 若要详细了解 Excel 中的自定义函数，请继续阅读以下文章：

> [!div class="nextstepaction"]
> [在 Excel 中创建自定义函数](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>法律信息

[IEX](https://iextrading.com/developer/) 免费提供的数据。 查看 [IEX 使用条款](https://iextrading.com/api-exhibit-a/)。 Microsoft 在本教程中使用的 IEX API 仅供教学使用。
