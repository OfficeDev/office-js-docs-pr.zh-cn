Office JavaScript API 包含两种不同的模型：

- **应用程序特定的** API 提供了强类型对象，它可用于与特定 Office 应用程序的本机对象进行交互。 例如，可使用 Excel JavaScript API 来访问工作表、区域、表格和图表等。 应用程序特定的 API 当前可用于以下 Office 应用程序。

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)
    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
    - [PowerPoint](../reference/overview/powerpoint-add-ins-reference-overview.md)
    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    此 API 模型使用的是[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，你可用它在你发送给 Office 应用程序的每个请求中指定多个操作。 通过以这种方式进行批处理操作，可大幅提升网页版 Office 应用程序中的加载项的性能。 应用程序特定的 API 是随 Office 2016 引入的，不可用于与 Office 2013 进行交互。

    > [!NOTE]
    > 还有一个应用程序特定的 [Visio](../reference/overview/visio-javascript-reference-overview.md) API，但它只能在 SharePoint Online 页面中用于与已嵌入到页面中的 Visio 图表进行交互。 Visio 不支持 Office Web 加载项。

    请参阅 [使用特定于应用程序的 API 模型](../develop/application-specific-api-model.md)，以了解有关此 API 模型的详细信息。

- **通用** API 可用于访问在多种类型的 Office 应用程序中都很常见的 UI、对话框和客户端设置等功能。 此 API 模型使用的是[回调](https://developer.mozilla.org/docs/Glossary/Callback_function)，这样,你在发送给 Office 应用程序的每个请求中只能指定一个操作。 通用 API 是随 Office 2013 引入的，可用于与 Office 2013 或更高版本进行交互。 要详细了解通用 API 对象模型（其中包括用于与 Outlook、PowerPoint 和 Project 交互的 API），请参阅[常见 JavaScript API 对象模型](../develop/office-javascript-api-object-model.md)。

> [!NOTE]
>没有 [共享运行时的](../testing/runtimes.md#shared-runtime) 自定义函数在 [仅限 JavaScript 的运行时中运行，该运行时](../testing/runtimes.md#javascript-only-runtime) 可确定计算执行的优先级。 这些函数使用略有不同的编程模型。
