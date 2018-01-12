# <a name="build-your-first-onenote-add-in"></a>生成你的第一个 OneNote 外接程序

本文介绍生成可将一些文本添加到 OneNote 页面的简单任务窗格外接程序的步骤。

下图显示将创建的外接程序。

   ![构建自此演练的 OneNote 外接程序](../../images/onenote-first-add-in.png)

<a name="setup"></a>
## <a name="step-1-set-up-your-dev-environment-and-create-an-add-in-project"></a>步骤 1：设置开发环境并创建外接程序项目
按照说明 [使用任何编辑器创建 Office 外接程序](../get-started/create-an-office-add-in-using-any-editor.md)，安装必需的系统必备组件并运行 Office Yeoman 生成器以创建新的外接程序项目。下表列出了要在 Yeoman 生成器中进行选择的项目属性。

| 选项 | 值 |
|:------|:------|
| 新建子文件夹 | （接受默认值） |
| 外接程序名称 | OneNote 外接程序 |
| 支持的 Office 应用程序 | （选择 OneNote） |
| 新建外接程序 | 是，我想要新建外接程序 |
| 添加 [TypeScript](https://www.typescriptlang.org/) | 否 |
| 选择框架 | Jquery |

<a name="develop"></a>
## <a name="step-2-modify-the-add-in"></a>步骤 2：修改外接程序
可以使用任何文本编辑器或 IDE 编辑外接程序文件。如果尚未尝试过 Visual Studio 代码，可以在 Linux、Mac OSX 和 Windows 上[免费下载](https://code.visualstudio.com/)。

1 - 打开项目目录中的 **index.html**。 

2 - 用以下代码替换 `<main>` 元素。这将添加使用 [Office UI Fabric 组件](http://dev.office.com/fabric/components)的文本区域和按钮。

```html
<main class="ms-welcome__main">
   <br />
   <p class="ms-font-l">Enter content below</p>
   <div class="ms-TextField ms-TextField--placeholder">
       <textarea id="textBox" rows="5"></textarea>
   </div>
   <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
        <span class="ms-Button-label">Add Outline</span>
        <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
        <span class="ms-Button-description">Adds the content above to the current page.</span>
    </button>
</main>
```

3 - 打开项目目录中的 **app.js**（或 app.ts，如果使用的是 TypeScript）。编辑 **Office.initialize** 函数，向“**添加边框**”按钮添加单击事件，如下所示。

```js
// The initialize function is run each time the page is loaded.
Office.initialize = function (reason) {
   $(document).ready(function () {
       app.initialize();
       
       // Set up event handler for the UI.
       $('#addOutline').click(addOutlineToPage);
   });
};
```
 
4 - 用以下 **addOutlineToPage** 方法替换 **run** 方法。这将从文本区域中获取内容，并将其添加至页面。

```js
// Add the contents of the text area to the page.
function addOutlineToPage() {        
   OneNote.run(function (context) {
      var html = '<p>' + $('#textBox').val() + '</p>';
      
       // Get the current page.
       var page = context.application.getActivePage();
       
       // Queue a command to load the page with the title property.             
       page.load('title'); 
       
       // Add an outline with the specified HTML to the page.
       var outline = page.addOutline(40, 90, html);
       
       // Run the queued commands, and return a promise to indicate task completion.
       return context.sync()
           .then(function() {
               console.log('Added outline to page ' + page.title);
           })
           .catch(function(error) {
               app.showNotification("Error: " + error); 
               console.log("Error: " + error); 
               if (error instanceof OfficeExtension.Error) { 
                   console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
               } 
           }); 
       });
}
```

<a name="test"></a>
## <a name="step-3-test-the-add-in-on-onenote-online"></a>步骤 3：在 OneNote Online 上测试外接程序
1 - 启动 HTTPS 服务器。  

  a.打开 **cmd** 提示符/终端，然后转到外接程序项目文件夹。 
  
  b.运行命令，如以下所示。

  ```
  C:\your-local-path\onenote add-in\> npm start
  ```

2 - 安装自签名证书作为受信任的证书。对于所有用 Office Yeoman 生成器创建的外接程序项目，只需在计算机上执行一次此操作。有关详细信息，请参阅[添加自签名证书作为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。

3 - 转到 [OneNote Online](https://www.onenote.com/notebooks)，然后打开一个笔记本。

4 - 选择“**插入 > Office 外接程序**”。该操作将打开 Office 外接程序对话框。

  -如果使用消费者帐户登录，请选择“**我的外接程序**”选项卡，然后选择“**上载我的外接程序**”。
  
  -如果使用工作或学校帐户登录，请选择“**我的组织**”选项卡，然后选择“**上载我的外接程序**”。 
  
  下图显示消费者笔记本的“**我的外接程序**”选项卡。

  ![显示“我的外接程序”选项卡的 Office 外接程序对话框](../../images/onenote-office-add-ins-dialog.png)

5 - 在“上载外接程序”对话框中，转到项目文件夹中的 **onenote-add-in-manifest.xml**，然后选择“**上载**”。测试时，清单文件会存储在浏览器的本地存储中。

6 - 该外接程序在 OneNote 页旁的 iFrame 中打开。在文本区域中输入一些文本，然后选择“**添加边框**”。您输入的文本将添加至页面。 

## <a name="troubleshooting-and-tips"></a>故障排除和提示
-可以使用浏览器的开发者工具调试外接程序。在 Internet Explorer 或 Chrome 中使用 Gulp Web 服务器并进行调试时，可以本地保存更改，然后仅刷新外接程序的 iFrame。

-检查 OneNote 对象时，目前可用的属性显示实际值。需要加载的属性显示“*未定义*”。展开 `_proto_` 节点以查看在对象上被定义但未加载的属性。

![在调试程序中上载 OneNote 对象](../../images/onenote-debug.png)

-如果你的外接程序使用任何 HTTP 资源，则需启用浏览器中的混合内容。生产外接程序应仅使用安全 HTTPS 资源。

-任务窗格外接程序可以从任何位置打开，但内容外接程序只能在常规页面内容（即不在标题、图像、IFrame 等中）内插入。 

## <a name="additional-resources"></a>其他资源

- [OneNote JavaScript API 编程概述](onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API 参考](../../reference/onenote/onenote-add-ins-javascript-reference.md)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 外接程序平台概述](https://dev.office.com/docs/add-ins/overview/office-add-ins)
