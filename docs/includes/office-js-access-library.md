可通过 Office JS 内容交付网络 (CDN) 访问 Office JavaScript API 库：`https://appsforoffice.microsoft.com/lib/1/hosted/Office.js` 要在任何加载项的网页中使用 Office JavaScript API，必须在页面的 `<head>` 标记中的 `<script>` 标记内引用 CDN。

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> 要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。

要详细了解如何访问 Office JavaScript API 库（包括如何获取 IntelliSense），请参阅[通过 Office JavaScript API 的内容交付网络 (CDN) 引用该库](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。