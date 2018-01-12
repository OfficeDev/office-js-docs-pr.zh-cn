
# <a name="guidelines-for-creating-labs-for-mix-using-labsjs"></a>使用 LabsJS 创建 Mix 实验室的准则



LabsJS 库 (labs.js) 支持编写与 Office Mix 集成的专门 Office 外接程序（称为实验室）。然后 Office Mix 使用 Microsoft PowerPoint 呈现实验室。我们将这些组件称为"实验室"，但我们应该明确我们创建的是特殊 Office 外接程序，即 Office Mix 外接程序。

LabsJS 内容提供了指导和示例，有助您实施 labs.js JavaScript API。此库在 [适用于 Office 的 JavaScript API](../../../reference/javascript-api-for-office.md) (Office.js) 的基础上构建，并提供了针对 Office Mix 中嵌入的外接程序优化的抽象层。


## <a name="general-guidelines"></a>通用准则


下面是在使用 LabJS API 编写外接程序时有帮助的一些通用准则。


### <a name="scripts"></a>脚本

因为 labs.js 库是 office.js 上的抽象层，因此依赖于 office.js，office.js 和 labs.js 库文件都必须包含在您的开发项目中。 

你可以引用此处的 office.js 库：`<script src="https://sforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>`。

labs.js 库附带 LabsJS SDK。或者，你可以引用 CDN <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code> 上的 labs.js 库。请注意，你的实验室的生产版本必须引用 CDN 上存储的版本。


 >**注意**：除了 JavaScript 文件 (labs-1.0.4.js) 以外，我们还提供实验室 API 的 TypeScript 定义文件 (labs-1.0.4.d.ts)。定义文件针对 TypeScript 版本 0.9.1.1 而构建。


### <a name="callbacks-and-error-handling"></a>回调和错误处理

labs.js API 中的几种方法将异步操作。对于这些操作，API 采用标准回调界面  **ILabCallback**。 


```js
function(err, result) {
}
```

回调方法采用两个参数： _err_ 和 _result_。 _err_ 字段仍为 **null**，除非存在错误。 _result_ 字段将返回操作的结果。

回调操作永远不会立即触发，即使结果立即可用。相反，它会在 JavaScript 事件循环单独执行时触发（通过调用  **setTimeout**）。通过采用此回调定义，可以将 labs.js 轻松地与您选择的承诺 API 集成。例如，您可以使用简单的转换方法替换这些回调的 jQuery 承诺，如以下示例中所示。




```js
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### <a name="lab-host-and-defaultlabhost"></a>实验室主机和 DefaultLabHost

实验室主机 ( **ILabHost**) 是支持实验室开发的基础驱动程序。默认情况下，它设置为与 office.js 集成的主机。

出于测试目的，并且为了在 labhost.html 内运行实验室，您需要切换到在模拟环境中工作的主机。以下代码示例演示如何使用查询参数执行此操作。或者，您可以将  **DefaultHostBuilder** 更改为将实验室外接程序与其他平台相集成。




```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### <a name="initialization"></a>初始化

初始化可在实验室及其主机之间建立通信通道。通过调用以下方法初始化您的实验室。


```js
Labs.connect((err, connectionResponse) => {});
```

初始化之后，您可以调用 labs.js API 的其他方法。 _connectionResponse_ 参数包含关于主机和用户的信息以及其他与连接相关的信息。有关返回值的详细信息，请参阅 [Labs.Core.IConnectionResponse](../../../reference/office-mix/labs.core.iconnectionresponse.md)。


### <a name="time-format"></a>时间格式

Labs.js 将数字存储为自 UTC 时间 1970 年 1 月 1 日以来经过的毫秒数。这与 JavaScript [Date 对象](http://msdn.microsoft.com/en-us/library/ie/cd9w2te4%28v=vs.94%29.aspx)的日期格式匹配。


### <a name="timeline"></a>时间线

实验室还可以与课程播放器时间线交互。时间线允许实验室通知课程播放器前进到下一张幻灯片。可通过调用  **Labs.getTimeline** 方法来检索时间线对象。


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="handling-events"></a>处理事件


LabsJS 事件 API 可跟踪实验室特定事件，并使您能够添加事件处理程序，以便对事件做出响应或采取措施。事件方法位于  **EventTypes** 对象上，其中三个为： **ModeChanged**、 **Activate** 和 **Deactivate**。 


### <a name="mode-change"></a>模式更改

当特定实验室从编辑模式更改为查看模式时， **ModeChanged** 事件将触发。在 PowerPoint 编辑模式下查看实验室时，编辑模式可见。PowerPoint 呈现幻灯片放映或实验室显示在 Office Mix 课程播放器中时，查看模式可见。查看模式应始终显示用户在使用实验室时所看到的内容。编辑模式允许用户配置实验室。

传递到回调的  **ModeChangedEventData** 对象中的数据包含有关当前模式的信息。以下代码演示如何使用 **ModeChanged** 事件。




```js
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### <a name="activate"></a>激活

当实验室所在的 PowerPoint 幻灯片在课程播放器中变为活动状态时， **activate** 事件将触发。


```js
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### <a name="deactivate"></a>停用

当实验室所在的 PowerPoint 幻灯片不再是活动幻灯片时， **deactivate** 事件将触发。


```js
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### <a name="timeline"></a>时间线

实验室还可以与课程播放器时间线交互。时间线允许实验室通知课程播放器前进到下一张幻灯片。可通过调用  **Labs.getTimeline** 方法来检索时间线对象。


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="additional-resources"></a>其他资源



- [Office Mix 外接程序](../../powerpoint/office-mix/office-mix-add-ins.md)
    
