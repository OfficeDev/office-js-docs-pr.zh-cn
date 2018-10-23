# <a name="tutorial-create-custom-functions-in-excel"></a>教程：在 Excel 中创建自定义函数

## <a name="introduction"></a>简介

自定义函数使你可以通过在 JavaScript 中定义这些函数作为加载项的一部分，将新函数添加到 Excel。然后，用户可以像使用 Excel 中的其他本机函数一样访问自定义函数，如 `SUM()`。你可以创建自定义函数，以执行简单任务，例如自定义计算或更复杂的任务，如将来自 Web 的实时数据l以流式处理方法插入工作表。

在此教程中，你将：
> [!div class="checklist"]
> * 使用 Yo Office 生成器创建自定义函数项目
> * 使用预建的自定义函数执行简单计算
> * 创建从 Web 请求数据的自定义函数
> * 创建以流式处理方法处理来自 Web 的实时数据的自定义函数

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>先决条件

* [Node.js 和 npm](https://nodejs.org/en/)

* [Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）

* [Yeoman](http://yeoman.io/) 和 [Yo Office 生成器](https://www.npmjs.com/package/generator-office)的最新版本。若要全局安装这些工具，请通过命令提示符处运行以下命令：

    ```bash
    npm install -g yo generator-office
    ```

* Excel for Windows（内部版本号 10827 或更高版本）或 Excel Online

* 加入 [Office  预览体验计划](https://products.office.com/office-insider)（**预览体验**级别 — 以前称为 "Insider Fast"）

## <a name="create-a-custom-functions-project"></a>创建自定义函数项目

你将通过使用 Yo Office 生成器创建自定义函数项目所需的文件来开始此教程。

1. 运行下面的命令，然后回答如下所示的提示问题。

    ```bash
    yo office
    ```

    * 选择项目类型： `Excel Custom Functions Add-in project (...)`
    * 选择脚本类型： `JavaScript`
    * 要如何命名加载项? `stock-ticker`

    ![自定义函数的 Yo Office bash 提示](../images/yo-office-cfs-stock-ticker-3.png)

    完成向导后，生成器会创建项目文件，并安装支持的 Node 组件。项目文件来自 [Excel 自定义函数](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub 存储库。

2. 导航到项目文件夹。

    ```bash
    cd stock-ticker
    ```

3. 启动本地 Web 服务器。

    * 如果你将使用 Excel for Windows 测试自定义函数，请运行以下命令以启动本地 Web 服务器、启动 Excel 和旁加载加载项：

        ```bash
        npm start
        ```

    * 如果你将使用 Excel Online 测试自定义函数，请运行以下命令以启动本地 Web 服务器： 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a>试用预建的自定义函数

使用 Yo Office 生成器创建的自定义函数项目包含一些在 **src/customfunction.js** 文件内定义的预建的自定义函数。在项目根目录中的 **Manifest.xml** 文件指定所有自定义函数属于 `CONTOSO` 命名空间。

在你可以使用任何预建的自定义函数之前，必须在 Excel 中注册自定义函数加载项。通过完成适用于你将在本教程中使用的平台的步骤来实现。

* 如果你将使用 Excel for Windows 测试自定义函数：

    1. 在 Excel 中，依次选择**插入**选项卡和位于**我的加载项**右侧的向下箭头。带突出显示“我的加载项”箭头的 Excel for Windows 中的 ![“插入”功能区](../images/excel-cf-register-add-in-1b.png)

    2. 在可用加载项列表中，找到**开发人员加载项**一节，然后选择 **Excel 自定义函数**加载项对其进行注册。![“我的加载项”列表中突出显示的带“Excel 自定义函数”加载项的 Excel for Windows 中的“插入”功能区](../images/excel-cf-register-add-in-2.png)

* 如果你将使用 Excel Online 测试自定义函数： 

    1. 在 Excel Online 中，依次选择**插入**选项卡和**加载项**。带突出显示“我的加载项”图标的 Excel Online 中的 ![“插入”功能区](../images/excel-cf-online-register-add-in-1.png)

    2. 依次选择**管理我的加载项**和**上传我的加载项**。 

    3. 选择**浏览...** 并导航到 Yo Office 生成器创建的项目的根目录。 

    4. 依次选择 **manifest.xml** 文件、**打开**和**上传**。

此时，在 Excel 内加载了项目中的预建自定义函数并可用。通过在 Excel 中完成以下步骤来试用 `ADD` 自定义函数：

1. 在单元格内，键入 **= CONTOSO**。请注意，自动完成菜单显示了 `CONTOSO` 命名空间中的所有函数列表。

2. 通过在单元格中指定下列值并按输入与作为输入参数的数字 `10` 和 `200` 一起来运行 `CONTOSO.ADD` 函数：

    ```
    =CONTOSO.ADD(10,200)
    ```

 `ADD` 自定义函数计算指定作为输入参数的两个数字的总和。在按 enter 后，键入 `=CONTOSO.ADD(10,200)` 应在单元格中生成结果 **210**。

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>创建从 Web 请求数据的自定义函数

要是需要一个可以通过 API 请求股票价格并在工作表的单元格中显示结果的函数会怎么样呢？设计自定义函数，以便可以轻松地从 Web 异步请求数据。

完成以下步骤以创建一个名为 `stockPrice` 的自定义函数，可接受股票代码（例如，**MSFT**），并返回该股票的价格。此自定义函数使用 IEX 交易 API，这是免费的，不需要身份验证。

1. 在 Yo Office 生成器创建的**股票代码**项目中，查找文件 **src/customfunctions.js** 并在代码编辑器中将其打开。

2. 将以下代码添加到 **customfunctions.js**，并保存文件。

    ```js
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

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. 在 Excel 可以使此新函数可供最终用户使用之前，必须指定介绍此函数的元数据。在 Yo Office 生成器创建的**股票代码** 项目中，查找文件 **config/customfunctions.json** 并在代码编辑器中将其打开。将以下对象添加到 **config/customfunctions.json** 文件内的 `functions` 数组，并保存文件。

    此 JSON 介绍了 `stockPrice` 函数。

    ```json
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. 必须在 Excel 中重新注册加载项，以便新函数可供最终用户使用。完成适用于你将在本教程中使用的平台的步骤。

    * 如果正在使用 Excel for Windows：

        1. 关闭 Excel，然后重新打开 Excel。

        2. 在 Excel 中，依次选择**插入**选项卡和位于**我的加载项**右侧的向下箭头。带突出显示“我的加载项”箭头的 Excel for Windows 中的 ![“插入”功能区](../images/excel-cf-register-add-in-1b.png)

        1. 在可用加载项列表中，找到**开发人员加载项**一节，然后选择 **Excel 自定义函数**加载项对其进行注册。![“我的加载项”列表中突出显示的带“Excel 自定义函数”加载项的 Excel for Windows 中的“插入”功能区](../images/excel-cf-register-add-in-2.png)

    * 如果正在使用 Excel Online： 

        1. 在 Excel Online 中，依次选择**插入**选项卡和**加载项**。带突出显示“我的加载项”图标的 Excel Online 中的 ![“插入”功能区](../images/excel-cf-online-register-add-in-1.png)

        2. 依次选择**管理我的加载项**和**上传我的加载项**。 

        3. 选择**浏览...** 并导航到 Yo Office 生成器创建的项目的根目录。 

        4. 依次选择 **manifest.xml** 文件、**打开**和**上传**。

5. 现在，让我们试用新函数。在单元格 **B1**中，键入文本 `=CONTOSO.STOCKPRICE("MSFT")` 并按 enter。应该看到单元格 **B1** 中的结果是针对 Microsoft 股票中的一支的当前股票价格。

## <a name="create-a-streaming-asynchronous-custom-function"></a>创建流式异步自定义函数

刚创建的 `stockPrice` 函数返回特定时刻的股票价格，但股票价格总是在变化。让我们一起创建一个通过 API 以流式处理传输数据以获取股票价格实时更新的自定义函数。

完成以下步骤以创建一个名为 `stockPriceStream` 自定义函数，请求指定每 1000 毫秒的价格股票（前提是在上一个请求已完成）。在进行的初始请求时，你可能会看到占位符值 **#GETTING_DATA** 单元格，函数正在这里调用。该函数返回一个值后，单元格中的那个值将替换 **#GETTING_DATA** 。

1. 在 Yo Office 生成器创建的**股票代码**项目中，将以下代码添加到 **src/customfunctions.js** 并保存文件。

    ```js
    function stockPriceStream(ticker, handler) {
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
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. 在 Excel 可以使此新函数可供最终用户使用之前，必须指定介绍此函数的元数据。在 Yo Office 生成器创建的**股票代码** 项目中，将以下对象添加到 **config/customfunctions.json** 文件内的 `functions` 数组，并保存文件。

    此 JSON 介绍了 `stockPriceStream` 函数。对于任何流式函数，`stream` 属性和 `cancelable` 属性必须设置为 `options` 对象内的 `true`，如此代码示例中所示。

    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

3. 必须在 Excel 中重新注册加载项，以便新函数可供最终用户使用。完成适用于你将在本教程中使用的平台的步骤。

    * 如果正在使用 Excel for Windows：

        1. 关闭 Excel，然后重新打开 Excel。
        
        2. 在 Excel 中，依次选择**插入**选项卡和位于**我的加载项**右侧的向下箭头。带突出显示“我的加载项”箭头的 Excel for Windows 中的 ![“插入”功能区](../images/excel-cf-register-add-in-1b.png)

        3. 在可用加载项列表中，找到**开发人员加载项**一节，然后选择 **Excel 自定义函数**加载项对其进行注册。![“我的加载项”列表中突出显示的带“Excel 自定义函数”加载项的 Excel for Windows 中的“插入”功能区](../images/excel-cf-register-add-in-2.png)

    * 如果正在使用 Excel Online： 

        1. 在 Excel Online 中，依次选择**插入**选项卡和**加载项**。带突出显示“我的加载项”图标的 Excel Online 中的 ![“插入”功能区](../images/excel-cf-online-register-add-in-1.png)

        2. 依次选择**管理我的加载项**和**上传我的加载项**。 

        3. 选择**浏览...** 并导航到 Yo Office 生成器创建的项目的根目录。 

        4. 依次选择 **manifest.xml** 文件、**打开**和**上传**。

4. 现在，让我们试用新函数。在单元格 **C1**中，键入文本 `=CONTOSO.STOCKPRICESTREAM("MSFT")` 并按 enter。假设股票市场处于打开状态，你应看到 **C1** 单元格中的结果不断地更新，以反映 Microsoft 股票中的一支的实时价格。

## <a name="next-steps"></a>后续步骤

在本教程中，你创建了一个新的自定义函数项目，尝试了预建的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 以流式处理方法传输实时数据的自定义函数。要了解在 Excel 中自定义函数的详细信息，请继续查看以下文章： 

> [!div class="nextstepaction"]
> [在 Excel 中创建自定义函数](../excel/custom-functions-overview.md)

## <a name="legal-information"></a>法律信息

由 [IEX](https://iextrading.com/developer/) 免费提供的数据。查看 [IEX 的使用条款](https://iextrading.com/api-exhibit-a/)。本教程中 Microsoft 使用的 IEX API 只用于教学目的。
