Outlook 外接程序主要使用通过[邮箱](/javascript/api/outlook/office.mailbox)对象公开的 api。 要访问专用于 Outlook 外接程序的对象和成员（例如 [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) 对象），可以使用 [Context](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) 对象的 **mailbox** 属性访问 **Mailbox** 对象，如下面的代码行所示。

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

另外，Outlook 外接程序可以使用以下对象：

-  **Office** 对象：用于初始化。

-  **Context** 对象：用于访问内容和显示语言属性。

-  **RoamingSettings** 对象：用于将 Outlook 加载项专用自定义设置保存到安装了加载项的用户邮箱。

有关使用 Outlook JavaScript API 的信息，请参阅[outlook 外接程序](../outlook/outlook-add-ins-overview.md)。