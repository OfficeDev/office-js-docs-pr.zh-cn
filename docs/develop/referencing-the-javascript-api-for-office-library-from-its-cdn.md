
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>从 适用于 Office 的 JavaScript API 的内容传送网络 (CDN) 引用 适用于 Office 的 JavaScript API 库


[适用于 Office 的 JavaScript API](http://dev.office.com/reference/add-ins/javascript-api-for-office) 库包含 Office.js 文件和关联的主机应用程序专有 .js 文件（如 Excel-15.js 和 Outlook-15.js）。 


引用 API 的最简单方法是，通过向页面的 `<head>` 标记添加以下 `<script>` 来使用我们的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

在 CDN URL 中，`office.js` 前面的 `/1/` 指定 Office.js 第 1 版中的最新增量版本。由于适用于 Office 的 JavaScript API 保留向后兼容性，因此最新版本将继续支持之前在第 1 版中引入的 API 成员。如果需要升级现有项目，请参阅[更新适用于 Office 的 JavaScript API 的版本和清单架构文件](update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果你计划从 Office 应用商店发布 Office 外接程序，必须使用此 CDN 引用。本地引用仅适用于内部、开发和调试应用场景。

> **重要说明：**开发任何 Office 主机应用程序的外接程序时，请从页面的 `<head>` 部分引用适用于 Office 的 JavaScript API。这样可确保 API 先于所有正文元素完全初始化。Office 主机要求外接程序在激活后的 5 秒内进行初始化。如果外接程序未在此阈值内激活，则会被声明为无响应，并且用户会看到错误消息。       

## <a name="additional-resources"></a>其他资源



- [了解适用于 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)    
- [适用于 Office 的 JavaScript API](http://dev.office.com/reference/add-ins/javascript-api-for-office)
    
