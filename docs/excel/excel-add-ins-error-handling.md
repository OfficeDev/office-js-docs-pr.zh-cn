# <a name="error-handling"></a>错误处理

使用 Excel JavaScript API 生成加载项时，请务必加入错误处理逻辑，以便解决运行时错误。 由于 API 的异步特性，此操作非常关键。

**注意**：有关 **sync()** 方法和 Excel JavaScript API 异步性的详细信息，请参阅 [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)。

## <a name="best-practices"></a>最佳做法

通过本文档中的代码示例，你会注意到每次调用 `Excel.run` 时，都会带上 `catch` 语句，以便捕获 `Excel.run` 内出现的任何错误。 建议在使用 Excel JavaScript API 生成加载项时使用相同模式。

```js
Excel.run(function (context) { 
  
  // Excel JavaScript API calls here

  // Await the completion of context.sync() before continuing.
  return context.sync()
    .then(function () {
      console.log("Finished!");
    })
}).catch(errorHandlerFunction);     
```

## <a name="api-errors"></a>API 错误 

当 Excel JavaScript API 请求无法成功运行时，API 将返回错误对象，其中包含以下属性： 

- **代码**：错误消息的 `code` 属性包含一个字符串，它属于 `OfficeExtension.ErrorCodes` 或 `Excel.ErrorCodes` 列表的一部分。 例如，错误代码“InvalidReference”指示引用对于指定操作无效。 错误代码尚未本地化。 

- **消息**：错误消息的 `message` 属性包含本地化字符串中的错误摘要。 错误消息并非针对最终用户的使用情况；应使用错误代码和相应的业务逻辑来确定加载项显示给最终用户的错误消息。

- **debugInfo**：出现此信息时，错误消息的 `debugInfo` 属性将提供其他信息，帮助理解错误根本原因。 

**注意**：如果使用 `console.log()` 将错误消息打印到控制台，那么这些消息将只能在服务器上可见。 最终用户不会在加载项任务窗格或主机应用程序的任何位置看到这些错误消息。

## <a name="additional-resources"></a>其他资源

- [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)
- [OfficeExtension.Error 对象（Excel JavaScript API）](http://dev.office.com/reference/add-ins/excel/error)
