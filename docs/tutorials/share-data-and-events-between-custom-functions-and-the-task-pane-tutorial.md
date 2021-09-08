---
title: 教程：Microsoft Excel自定义函数和任务窗格之间共享数据和事件
description: 学习如何在Microsoft Excel中的自定义函数和任务窗格之间共享数据和事件。
ms.date: 08/04/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: fe304ff0eef4b707cb0b90a2c38c60aff921279d
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939102"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>教程：Microsoft Excel自定义函数和任务窗格之间共享数据和事件

你可配置 Excel 加载项以使用共享运行时。 这样就可以共享全局数据，或者发送任务窗格和自定义功能之间的事件。

对于大多数自定义函数方案，建议使用共享运行时，除非有特定的理由使用非任务窗格（UI-less）自定义函数。

本教程假设你已经熟悉使用Yo Office生成器来创建插件项目。 如果尚未完成[Excel 自定义函数教程](excel-tutorial-create-custom-functions.md)，请考虑完成它。

## <a name="create-the-add-in-project"></a>创建加载项项目

使用 Yeoman 生成器创建 Excel 加载项项目。运行以下命令，然后使用以下答案回答提示。

```command line
yo office
```

- 选择项目类型： **Excel 自定义函数加载项项目**
- 选择脚本类型： **JavaScript**
- 你想要如何命名加载项？ **我的 Office 加载项**

![显示命令行界面中 Yeoman 生成器的提示和回答的屏幕截图。](../images/yo-office-excel-project.png)

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

## <a name="configure-the-manifest"></a>配置清单

1. 启动 Visual Studio Code 并打开“**我的 Office 加载项**”项目。
2. 打开 **manifest.xml** 文件。
3. 找到 `<VersionOverrides>` 部分并添加以下 `<Runtimes>` 部分。 生存期需要 **较长**，以便在关闭任务窗格时自定义函数仍可正常工作。

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

> [!NOTE]
> 如果加载启动项包括清单中的 `Runtimes` 元素，则无论 Windows 或 Microsoft 365 版本如何，都将使用 Internet Explorer 11。 有关详细信息，请参阅[运行时](../reference/manifest/runtimes.md)。

4. 在 `<Page>` 元素中，将源位置从 **Functions.Page.Url** 更改为 **ContosoAddin.Url**。

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. 在 `<DesktopFormFactor>` 部分中，将 **FunctionFile** 从 **Commands.Url** 更改为使用 **ContosoAddin.Url**。

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. 在 `<Action>` 部分中，将源位置从 **Taskpane.Url** 更改为 **ContosoAddin.Url**。

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. 为 **ContosoAddin.Url** 添加新的 **Url id**，它指向 **taskpane.html**。

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. 保存更改并重新生成项目。

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>共享自定义函数和任务窗格代码之间的状态

由于自定义函数在与任务窗格代码相同的上下文中运行，因此可以直接共享状态，无需使用 **Storage** 对象。 以下说明介绍了如何在自定义函数和任务窗格代码之间共享全局变量。

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>创建用于获取或存储共享状态的自定义函数

1. 在 Visual Studio Code 中，打开文件 **src/functions/functions.js**。
2. 在第 1 行，将以下代码插入到最顶部。 这将初始化一个名为 **sharedState** 的全局变量。

   ```js
   window.sharedState = "empty";
   ```

3. 添加以下代码，创建将值存储到 **sharedState** 变量的自定义函数。

   ```js
   /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
   function storeValue(sharedValue) {
     window.sharedState = sharedValue;
     return "value stored";
   }
   ```

4. 添加以下代码，创建获取 **sharedState** 变量的当前值的自定义函数。

   ```js
   /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
   function getValue() {
     return window.sharedState;
   }
   ```

5. 保存文件。

### <a name="create-task-pane-controls-to-work-with-global-data"></a>创建任务窗格控件以处理全局数据

1. 打开 **src/taskpane/taskpane.html** 文件。
2. 紧贴在结尾的 `</head>` 元素前，添加以下脚本元素。

   ```html
   <script src="functions.js"></script>
   ```

3. 关闭 `</main>` 元素后，添加以下 HTML。 该 HTML 创建两个用于获取或存储全局数据的文本框和按钮。

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
     <li>Select <strong>Get</strong> to display the value in the task pane.</li>
   </ol>
   <p>Store new value to shared state</p>
   <div>
     <input type="text" id="storeBox" />
     <button onclick="storeSharedValue()">Store</button>
   </div>

   <p>Get shared state value</p>
   <div>
     <input type="text" id="getBox" />
     <button onclick="getSharedValue()">Get</button>
   </div>
   ```

4. 在结尾的 `</body>` 元素前，添加以下脚本。 当用户想存储或获取全局数据时，此代码将处理按钮单击事件。

   ```js
   <script>
   function storeSharedValue() {
     let sharedValue = document.getElementById('storeBox').value;
     window.sharedState = sharedValue;
   }

   function getSharedValue() {
     document.getElementById('getBox').value = window.sharedState;
   }
   </script>
   ```

5. 保存文件。
6. 生成项目

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>在自定义函数和任务窗格之间尝试共享数据

- 使用以下命令启动项目。

  ```command line
  npm run start
  ```

Excel 启动后，可使用“任务窗格”按钮来存储或获取共享数据。 在自定义函数的单元格中输入 `=CONTOSO.GETVALUE()`，以检索相同的共享数据。 或使用 `=CONTOSO.STOREVALUE("new value")` 将共享数据更改为新值。

> [!NOTE]
> 如本文所示配置项目，可在自定义函数和任务窗格之间共享上下文。 通过自定义函数可以调用一些Office API。 更多细节请参见[从自定义函数调用Microsoft Excel APIs](../excel/call-excel-apis-from-custom-function.md)。
