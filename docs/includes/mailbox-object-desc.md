<span data-ttu-id="eae8a-101">Outlook 外接程序主要使用通过[邮箱](/javascript/api/outlook/office.mailbox)对象公开的 api。</span><span class="sxs-lookup"><span data-stu-id="eae8a-101">Outlook add-ins primarily use the APIs exposed through the [Mailbox](/javascript/api/outlook/office.mailbox) object.</span></span> <span data-ttu-id="eae8a-102">要访问专用于 Outlook 外接程序的对象和成员（例如 [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) 对象），可以使用 [Context](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) 对象的 **mailbox** 属性访问 **Mailbox** 对象，如下面的代码行所示。</span><span class="sxs-lookup"><span data-stu-id="eae8a-102">To access the objects and members specifically for use in Outlook add-ins, such as the [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) object, you use the [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.</span></span>

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

<span data-ttu-id="eae8a-103">另外，Outlook 外接程序可以使用以下对象：</span><span class="sxs-lookup"><span data-stu-id="eae8a-103">Additionally, Outlook add-ins can use the following objects:</span></span>

-  <span data-ttu-id="eae8a-104">**Office** 对象：用于初始化。</span><span class="sxs-lookup"><span data-stu-id="eae8a-104">**Office** object: for initialization.</span></span>

-  <span data-ttu-id="eae8a-105">**Context** 对象：用于访问内容和显示语言属性。</span><span class="sxs-lookup"><span data-stu-id="eae8a-105">**Context** object: for access to content and display language properties.</span></span>

-  <span data-ttu-id="eae8a-106">**RoamingSettings** 对象：用于将 Outlook 加载项专用自定义设置保存到安装了加载项的用户邮箱。</span><span class="sxs-lookup"><span data-stu-id="eae8a-106">**RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.</span></span>

<span data-ttu-id="eae8a-107">有关使用 Outlook JavaScript API 的信息，请参阅[outlook 外接程序](../outlook/outlook-add-ins-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="eae8a-107">For information about using the Outlook JavaScript API, see [Outlook add-ins](../outlook/outlook-add-ins-overview.md).</span></span>