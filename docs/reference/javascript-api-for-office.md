# <a name="javascript-api-for-office"></a>适用于 Office 的 JavaScript API

JavaScript API for Office 让您能够创建与 Office 主机应用程序中的对象模型交互的 Web 应用程序。 应用程序将引用 office.js 库，它是脚本加载程序。 office.js 库加载适用于运行加载项的 Office 应用程序的对象模型。 您可以使用以下 JavaScript 对象模型：

- **通用 API** - 与 **Office 2013** 一起引入的 API。 这是为**所有 Office 主机应用程序**加载的，并将加载项应用程序与 Office 客户端应用程序连接。 对象模型包含特定于 Office 客户端的 API，以及适用于多个 Office 客户端主机应用程序的 API。 所有内容都在**共享 API** 下。 

  **Outlook** 也使用通用 API 语法。 别名 Office 下的所有内容都包含可用于编写脚本的对象，这些脚本与Office 文档、工作表、演示文稿、邮件项和 Office 加载项项目中的内容进行交互。如果加载项的目的用于 Office 2013 及更高版本，则必须使用这些通用 API。 此对象模型使用回调。

- **特定于主机的 API** - 与 **Office 2016** 一起引入的 API。 此对象模型提供特定于主机的强类型对象，这些对象对应于使用 Office 客户端时所看到的熟悉对象，并代表未来的 Office JavaScript API。 特定于主机的 API 目前包括 Word JavaScript API 和 Excel JavaScript API。

## <a name="supported-host-applications"></a>支持的主机应用程序

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [共享 API](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> [PowerPoint 和 Project](requirement-sets/powerpoint-and-project-note.md) 支持加载项使用 JavaScript API 制作而成。 但是，它们当前没有特定于主机的 API。 您通过共享 API 与这些主机进行交互。

了解有关[支持的主机和其他要求](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins)的详细信息。

## <a name="open-api-specifications"></a>开放 API 规范

在我们设计和开发新的 API 以用于 Office 加载项时，我们将在[开放 API 规范](openspec.md)页面公开收集您的反馈。了解正处于准备阶段的新增功能，并提供您对我们的设计规范的宝贵意见。

## <a name="see-also"></a>另请参阅

- [Office 的 JavaScript API 参考](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)