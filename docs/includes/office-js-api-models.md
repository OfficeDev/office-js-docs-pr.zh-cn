Office JavaScript API 包含两种不同的模型：

- **主机特定的** API 提供了强类型对象，它可用于与特定 Office 应用程序的本机对象进行交互。 例如，可使用 Excel JavaScript API 来访问工作表、区域、表格和图表等。 主机特定的 API 当前可用于以下主机：

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)

    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)

    此 API 模型使用的是[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)，你可用它在你发送给 Office 主机的每个请求中指定多个操作。 通过此方式批量处理操作，可大幅提升 Web 应用程序上的 Office 中的性能。 主机特定的 API 是随 Office 2016 引入的，不可用于与 Office 2013 进行交互。

- **通用** API 可用于访问在多种类型的 Office 应用程序中都很常见的 UI、对话框和客户端设置等功能。 此 API 模型使用的是[回调](https://developer.mozilla.org/docs/Glossary/Callback_function)，其中你仅可在发送给 Office 主机的每个请求中指定一个操作。 通用 API 是随 Office 2013 引入的，可用于与 Office 2013 或更高版本进行交互。 要详细了解通用 API 对象模型（其中包括用于与 Outlook 和 PowerPoint 交互的 API），请参阅[常见 JavaScript API 对象模型](../develop/office-javascript-api-object-model.md)。

> [!NOTE]
> Excel 自定义函数在排列了计算执行优先级的唯一运行时中运行，因此使用的编程模型略有不同。 有关详细信息，请参阅[自定义函数体系结构](../excel/custom-functions-architecture.md)。