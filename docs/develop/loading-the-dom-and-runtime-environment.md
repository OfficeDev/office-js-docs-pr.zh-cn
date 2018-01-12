
# <a name="loading-the-dom-and-runtime-environment"></a>加载 DOM 和运行时环境



外接程序在运行自己的自定义逻辑前必须确保 DOM 和 Office 外接程序运行时环境都已加载。 

## <a name="startup-of-a-content-or-task-pane-add-in"></a>启动内容或任务窗格加载项

下图显示了在 Excel、PowerPoint、Project、Word 或 Access 中启动内容或任务窗格外接程序所涉及的事件流。

![启动内容/任务窗格外接程序时的事件流](../../images/off15appsdk_LoadingDOMAgaveRuntime.png)

启动内容/任务窗格外接程序时，将发生以下事件： 



1. 用户打开已包含加载项的文档或在文档中插入加载项。
    
2. Office 主机应用程序从 Office 商店、SharePoint 的加载项目录或者其源于的共享文件夹目录中读取加载项的 XML 清单。
    
3. Office 主机应用程序在浏览器控件中打开加载项的 HTML 页面。
    
    后面的两个步骤第 4 步和第 5 步以异步方式并行发生。因此，您的加载项代码必须在继续之前确保 DOM 和加载项运行时环境已加载完。
    
4. 浏览器控件加载 DOM 和 HTML 正文，并调用  **window.onload** 事件的事件处理程序。
    
5. Office 主机应用程序加载运行时环境，这将从内容分发网络 (CDN) 服务器中为 JavaScript 库文件下载并缓存 JavaScript API，然后为 [Office](../../reference/shared/office.initialize.md) 对象的 [initialize](../../reference/shared/office.md) 事件调用加载项的事件处理程序。
    
6. 当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。
    

## <a name="startup-of-an-outlook-add-in"></a>启动 Outlook 外接程序



下图显示了启动在台式机、平板电脑或智能手机上运行的 Outlook 外接程序所涉及的事件流。

![启动 Outlook 外接程序时的事件流](../../images/olowawecon15_LoadingDOMAgaveRuntime.png)

启动 Outlook 外接程序时，将发生以下事件： 



1. 当 Outlook 启动时，Outlook 读取已为用户的电子邮件帐户安装的 Outlook 外接程序的 XML 清单。
    
2. 用户选择 Outlook 中的一个项目。
    
3. 如果所选项目满足某个 Outlook 外接程序的激活条件，则 Outlook 将激活该外接程序，并使其按钮在 UI 中可见。
    
4. 如果用户单击该按钮以启动 Outlook 外接程序，Outlook 将在浏览器控件中打开 HTML 页面。下面两个步骤（步骤 5 和 6）并行发生。
    
5. 浏览器控件加载 DOM 和 HTML 正文，并调用  **onload** 事件的事件处理程序。
    
6. Outlook 调用加载项的 [Office](../../reference/shared/office.initialize.md) 对象的 [initialize](../../reference/shared/office.md) 事件的事件处理程序。
    
7. 当 DOM 和 HTML 正文加载完毕并且加载项完成初始化后，加载项的主函数就可以继续进行。
    

## <a name="checking-the-load-status"></a>检查加载状态


检查 DOM 和 运行时环境是否加载完毕的一种方式是使用 jQuery [.ready()](http://api.jquery.com/ready/) 函数： `$(document).ready()`。例如，以下  **initialize** 事件处理程序函数可确保在专门用于初始化外接程序的代码运行前先加载 DOM。随后， **initialize** 事件处理程序继续使用 [mailbox.item](../../reference/outlook/Office.context.mailbox.item.md) 属性获取 Outlook 中当前选定的项目，并调用外接程序的主函数 `initDialer`。


```js
Office.initialize = function () {
    // Checks for the DOM to load.
    $(document).ready(function () {
        // After the DOM is loaded, add-in-specific code can run.
        var mailbox = Office.context.mailbox;
        _Item = mailbox.item;
        initDialer();
    });
}
```

这种方法可在任何 Office 外接程序的  **initialize** 处理程序中使用。

电话拨号程序示例 Outlook 外接程序演示了略为不同的方法，此方法仅使用 JavaScript 检查这些相同条件。 

 **重要说明：**即使你的加载项无需执行初始化任务，你也必须至少加入一个如下最小的 **Office.initialize** 事件处理程序函数。




```js
Office.initialize = function () {
};
```

如果您无法加入  **Office.initialize** 事件处理程序，则启动加载项时可能会出错。此外，如果用户尝试将您的加载项与 Office Online Web 客户端（如 Excel Online、PowerPoint Online 或 Outlook Web App）结合使用，应用程序将无法运行。

如果您的加载项包括多个页，则在每次加载新页时，页面必须加入或调用  **Office.initialize** 事件处理程序。


## <a name="additional-resources"></a>其他资源



- [了解适用于 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
