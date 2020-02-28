<span data-ttu-id="60542-101">可通过 Office JS 内容交付网络 (CDN) 访问 Office JavaScript API 库：`https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`</span><span class="sxs-lookup"><span data-stu-id="60542-101">The Office JavaScript API library can be accessed via the Office JS content delivery network (CDN) at: `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`.</span></span> <span data-ttu-id="60542-102">要在任何加载项的网页中使用 Office JavaScript API，必须在页面的 `<head>` 标记中的 `<script>` 标记内引用 CDN。</span><span class="sxs-lookup"><span data-stu-id="60542-102">To use Office JavaScript APIs within any of your add-in's web pages, you must reference the CDN in a `<script>` tag in the `<head>` tag of the page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
</head>
```

> [!NOTE]
> <span data-ttu-id="60542-103">要使用预览版 API，请参考 CDN 上的 Office JavaScript API 库预览版：`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`。</span><span class="sxs-lookup"><span data-stu-id="60542-103">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

<span data-ttu-id="60542-104">要详细了解如何访问 Office JavaScript API 库（包括如何获取 IntelliSense），请参阅[通过 Office JavaScript API 的内容交付网络 (CDN) 引用该库](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md)。</span><span class="sxs-lookup"><span data-stu-id="60542-104">For more information about accessing the Office JavaScript API library, including how to get IntelliSense, see [Referencing the Office JavaScript API library from its content delivery network (CDN)](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>