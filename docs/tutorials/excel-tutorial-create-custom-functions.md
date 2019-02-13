---
title: Excel 自定义函数教程（预览）
description: 在本教程中，你将创建一个 Excel 外接程序，其中包含可执行计算、请求 Web 数据或流式传输 Web 数据的自定义函数。
ms.date: 01/08/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 4ac735e6fc19f13859d07df6cb3d2443e6dfe2fd
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982018"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a>教程：在 Excel 中创建自定义函数（预览）

用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。 Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。 可以创建自定义函数，以执行简单的任务（如计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。

在本教程中，你将：
> [!div class="checklist"]
> * 使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建自定义函数加载项。 
> * 使用预生成的自定义函数来执行简单计算。
> * 创建从 Web 获取数据的自定义函数。
> * 创建从 Web 传送实时数据的自定义函数。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a>先决条件

* [Node.js](https://nodejs.org/en/)（版本 8.0.0 或更高版本）

* [Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）

* 最新版本的 [Yeoman](https://yeoman.io/) 和[适用于 Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > 即便先前已安装 Yeoman 生成器，我们仍建议将包更新至最新的 npm 版本。

* Excel for Windows（64 位，版本 1810 或更高版本）或 Excel Online

* 加入 [Office 预览体验计划](https://products.office.com/office-insider)（**预览体验成员**级别 - 以前称为“预览体验成员 - 快”）

## <a name="create-a-custom-functions-project"></a>创建自定义函数项目

 首先，创建代码项目以构建自定义函数加载项。 [适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)将使用可供你试用的一些初始自定义函数来设置项目。

1. 运行下面的命令，再回答如下所示的提示问题。
    
    ```
    yo office
    ```
    
    * 选择项目类型：`Excel Custom Functions Add-in project (...)`
    * 选择脚本类型：`JavaScript`
    * 要如何命名加载项？ `stock-ticker`
    
    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/12-10-fork-cf-pic.jpg)
    
    Yeoman 生成器将创建项目文件并安装支持的 Node.js 组件。

2. 转到项目文件夹。
    
    ```
    cd stock-ticker
    ```

3. 信任运行此项目所需的自签名证书。 有关适用于 Windows 或 Mac 的详细说明，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。  

4. 生成项目。
    
    ```
    npm run build
    ```

5. 启动在 Node.js 中运行的本地 Web 服务器。 你可以在 Excel for Windows 或 Excel Online 中尝试使用自定义函数加载项。

# <a name="excel-for-windowstabexcel-windows"></a>[Excel for Windows](#tab/excel-windows)

运行以下命令。

```
npm run start
```

此命令将启动 Web 服务器，并将自定义函数加载项旁加载到 Excel for Windows 中。

> [!NOTE]
> 如果加载项未加载，请检查是否已正确完成步骤 3。 您还可以**[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** 来解决问题的外接程序的 XML 指令清单文件，以及任何安装或运行时的问题。 运行时日志记录写入`console.log`语句日志文件以帮助您查找和修复问题。

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

运行以下命令。

```
npm run start-web
```

此命令将启动 Web 服务器。 使用以下步骤来旁加载你的加载项。

<ol type="a">
   <li>在 Excel Online 中，依次选择“插入”<strong></strong>选项卡和“加载项”<strong></strong>。<br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li>选择“管理我的加载项”<strong></strong>，然后选择“上载我的加载项”<strong></strong>。</li> 
   <li>选择“浏览...”<strong></strong>，并导航到 Yeoman 生成器创建的项目的根目录。</li> 
   <li>依次选择文件“manifest.xml”<strong></strong>，“打开”<strong></strong>，然后选择“上载”<strong></strong>。</li>
</ol>

> [!NOTE]
> 如果加载项未加载，请检查是否已正确完成步骤 3。

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a>尝试预生成的自定义函数

你创建的自定义函数项目已经有两个预生成的自定义函数，名为 ADD 和 INCREMENT。 这些预生成的函数的代码位于 **src/customfunctions.js** 文件中。 **./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。 你将使用 CONTOSO 命名空间来访问 Excel 中的自定义函数。

接下来，通过完成以下步骤来尝试使用 `ADD` 自定义函数：

1. 在 Excel 中，转至任意单元格并输入 `=CONTOSO`。 请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。

2. 通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。

`ADD` 自定义函数将计算你提供的两个数字的总和，并返回结果 **210**。

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>创建从 Web 请求数据的自定义函数

集成来自 Web 的数据是通过自定义函数来扩展 Excel 的好方法。 接下来，你将创建一个名为 `stockPrice` 的自定义函数，该函数从 Web API 获取股票报价并将结果返回到工作表的单元格。 你将使用使用 IEX Trading API，该 API 是免费的，并且不需要身份验证。

1. 在 **stock-ticker** 项目中，找到文件 **src/customfunctions.js** 并在代码编辑器中打开它。

2. 在 **customfunctions.js** 中，找到 `increment` 函数并将以下代码添加到该函数后面。

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

> [!NOTE]
> In the January Insiders 1901 Build, there is a bug preventing fetch calls from executing which will result in #VALUE!.
> To workaround this please use the [XMLHTTPRequest API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime#requesting-external-data) to make the web request.

3. In **customfunctions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    `CustomFunctions.associate` 代码会将函数的 `id` 与 JavaScript 中的 `increment` 的函数地址相关联，以便 Excel 能够调用你的函数。

    在 Excel 能够使用你的自定义函数之前，你需要先使用元数据来描述它。 你需要先定义在 `associate` 方法中使用的 `id` 以及某些其他元数据。


4. 打开 **config/customfunctions.json** 文件。 将 JSON 对象添加到“函数”数组中，然后保存该文件。

    ```JSON
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
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    此 JSON 将描述 `stockPrice` 函数、其参数以及它返回的结果类型。

5. 在 Excel 中重新注册加载项，以便新函数可用。 

# <a name="excel-for-windowstabexcel-windows"></a>[Excel for Windows](#tab/excel-windows)

1. 关闭 Excel，然后重新打开 Excel。

2. 在 Excel 中，选择“插入”**** 选项卡，然后选择位于“我的加载项”**** 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)

3. 在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。
    ![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)

# <a name="excel-onlinetabexcel-online"></a>[Excel Online](#tab/excel-online)

1. 在 Excel Online 中，选择“插入”**** 选项卡，然后选择“加载项”****。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)

2. 选择“管理我的加载项”****，然后选择“上载我的加载项”****。 

3. 选择“浏览...”****，并导航到 Yeoman 生成器创建的项目的根目录。 

4. 依次选择文件“manifest.xml”****，“打开”****，然后选择“上载”****。

--- 

<ol start="6">
<li> 尝试使用新函数。 在单元格 <strong>B1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICE("MSFT")</strong>，然后按 Enter。 应看到单元格 <strong>B1</strong> 中的结果是 Microsoft 一股股票的当前股票价格。</li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a>创建流式处理异步自定义函数

`stockPrice` 函数将返回特定时刻的股票价格，但股票价格一直在变化。 接下来，将创建一个名为 `stockPriceStream` 的自定义函数，该函数每隔 1000 毫秒获取一次股票价格。

1. 在 **stock-ticker** 项目中，将以下代码添加到 **src/customfunctions.js** 并保存该文件。

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
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    在 Excel 能够使用你的自定义函数之前，你需要先使用元数据来描述它。
    
2. 在 **stock-ticker** 项目中，将以下对象添加到 **config/customfunctions.json** 文件中的 `functions` 数组，并保存该文件。
    
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
                "description": "stock symbol",
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

    此 JSON 说明了 `stockPriceStream` 函数。 对于任何流式处理函数，必须在 `options` 对象中将 `stream` 属性和 `cancelable` 属性设置为 `true`，如本代码示例所示。

3. 在 Excel 中重新注册加载项，以便新函数可用。

# <a name="excel-for-windowstabexcel-windows"></a>[Excel for Windows](#tab/excel-windows)

1. 关闭 Excel，然后重新打开 Excel。

2. 在 Excel 中，选择“插入”**** 选项卡，然后选择位于“我的加载项”**** 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)

3. 在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。
    ![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)

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

恭喜！ 你已经创建新的自定义函数项目，尝试了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 传送实时数据的自定义函数。 若要详细了解 Excel 中的自定义函数，请继续阅读以下文章：

> [!div class="nextstepaction"]
> [在 Excel 中创建自定义函数](../excel/custom-functions-overview.md)

### <a name="legal-information"></a>法律信息

[IEX](https://iextrading.com/developer/) 免费提供的数据。 查看 [IEX 使用条款](https://iextrading.com/api-exhibit-a/)。 Microsoft 在本教程中使用的 IEX API 仅供教学使用。


