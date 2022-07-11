Outlook 外接程序主要使用通过 [Mailbox](/javascript/api/outlook/office.mailbox) 对象公开的 API 的子集。 若要访问专门用于 Outlook 外接程序（如 [Item](/javascript/api/outlook/office.item) 对象）的对象和成员，请使用 **Context** 对象的 [邮箱](/javascript/api/office/office.context#office-office-context-mailbox-member)属性访问 **邮箱** 对象，如以下代码行所示。

```js
// Access the Item object.
const item = Office.context.mailbox.item;
```

此外，Outlook 加载项可以使用以下对象。

- **Office** 对象：用于初始化。

- **Context** 对象：用于访问内容和显示语言属性。

- **RoamingSettings** 对象：用于将 Outlook 加载项专用自定义设置保存到安装了加载项的用户邮箱。

有关在 Outlook 加载项中使用 JavaScript 的信息，请参阅 [Outlook 加载项](../outlook/outlook-add-ins-overview.md)。
