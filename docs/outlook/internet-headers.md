---
title: 获取和设置 internet 标头
description: 如何：在 Outlook 外接程序中获取和设置邮件的 internet 邮件头。
ms.date: 11/04/2019
localization_priority: Normal
ms.openlocfilehash: d7f38b7564683ce51ed0bd840480b4a8b2040bf6
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165998"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a><span data-ttu-id="ef1dc-103">在 Outlook 外接程序中获取和设置邮件的 internet 邮件头</span><span class="sxs-lookup"><span data-stu-id="ef1dc-103">Get and set internet headers on a message in an Outlook add-in</span></span>

## <a name="background"></a><span data-ttu-id="ef1dc-104">背景</span><span class="sxs-lookup"><span data-stu-id="ef1dc-104">Background</span></span>

<span data-ttu-id="ef1dc-105">Outlook 外接程序开发中的一个常见要求是，将与外接程序关联的自定义属性存储在不同的级别。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-105">A common requirement in Outlook add-ins development is to store custom properties associated with an add-in at different levels.</span></span> <span data-ttu-id="ef1dc-106">目前，自定义属性存储在项目或邮箱级别。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-106">At present, custom properties are stored at the item or mailbox level.</span></span>

- <span data-ttu-id="ef1dc-107">项目级别-适用于特定项目的属性，使用[CustomProperties](/javascript/api/outlook/office.customproperties)对象。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-107">Item level - For properties that apply to a specific item, use the [CustomProperties](/javascript/api/outlook/office.customproperties) object.</span></span> <span data-ttu-id="ef1dc-108">例如，存储与发送电子邮件的人员关联的客户代码。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-108">For example, store a customer code associated with the person who sent the email.</span></span>
- <span data-ttu-id="ef1dc-109">邮箱级别-对于应用于用户邮箱中的所有邮件项目的属性，使用[RoamingSettings](/javascript/api/outlook/office.roamingsettings)对象。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-109">Mailbox level - For properties that apply to all the mail items in the user's mailbox, use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object.</span></span> <span data-ttu-id="ef1dc-110">例如，存储用户首选项以按特定比例显示温度。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-110">For example, store a user's preference to show the temperature in a particular scale.</span></span>

<span data-ttu-id="ef1dc-111">在项目离开 Exchange 服务器后，这两种类型的属性都不会保留，因此电子邮件收件人无法获取项目上设置的任何属性。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-111">Both types of properties are not preserved after the item leaves the Exchange server so the email recipients can't get any properties set on the item.</span></span> <span data-ttu-id="ef1dc-112">因此，开发人员无法访问这些设置或其他 MIME 属性以实现更好的阅读方案。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-112">Therefore, developers can't access those settings or other MIME properties to enable better read scenarios.</span></span>

<span data-ttu-id="ef1dc-113">虽然有一种方法可以将 internet 标头设置为 EWS 请求，但在某些情况下，不能进行 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-113">While there's a way for you to set the internet headers through EWS requests, in some scenarios making an EWS request won't work.</span></span> <span data-ttu-id="ef1dc-114">例如，在 Outlook 桌面的撰写模式下，项目 id 在缓存模式下 `saveAsync` 不会同步。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-114">For example, in Compose mode on Outlook desktop, the item id isn't synced on `saveAsync` in cached mode.</span></span>

> [!TIP]
> <span data-ttu-id="ef1dc-115">请参阅[获取和设置 Outlook 外接程序的外接程序元数据](metadata-for-an-outlook-add-in.md)，以了解有关使用这些选项的详细信息。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-115">See [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md) to learn more about using these options.</span></span>

## <a name="purpose-of-the-internet-headers-api"></a><span data-ttu-id="ef1dc-116">Internet 标头 API 的用途</span><span class="sxs-lookup"><span data-stu-id="ef1dc-116">Purpose of the internet headers API</span></span>

<span data-ttu-id="ef1dc-117">在要求集1.8 中引入，internet 标头 Api 使开发人员能够：</span><span class="sxs-lookup"><span data-stu-id="ef1dc-117">Introduced in requirement set 1.8, the internet headers APIs enable developers to:</span></span>

- <span data-ttu-id="ef1dc-118">戳在所有客户端上保留 Exchange 后保留的电子邮件的信息。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-118">Stamp information on an email that persists after it leaves Exchange across all clients.</span></span>
- <span data-ttu-id="ef1dc-119">阅读有关在邮件读取应用场景中的所有客户端上的电子邮件保留后保留的电子邮件的信息。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-119">Read information on an email that persisted after the email left Exchange across all clients in mail read scenarios.</span></span>
- <span data-ttu-id="ef1dc-120">访问电子邮件的整个 MIME 标头。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-120">Access the entire MIME header of the email.</span></span>

## <a name="set-internet-headers-while-composing-a-message"></a><span data-ttu-id="ef1dc-121">在撰写邮件时设置 internet 邮件头</span><span class="sxs-lookup"><span data-stu-id="ef1dc-121">Set internet headers while composing a message</span></span>

<span data-ttu-id="ef1dc-122">尝试使用[internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders)属性来管理在撰写模式下放置在当前邮件上的自定义 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-122">Try using the [item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) property to manage the custom internet headers you place on the current message in Compose mode.</span></span>

### <a name="set-get-and-remove-custom-headers-example"></a><span data-ttu-id="ef1dc-123">设置、获取和删除自定义标头示例</span><span class="sxs-lookup"><span data-stu-id="ef1dc-123">Set, get, and remove custom headers example</span></span>

<span data-ttu-id="ef1dc-124">下面的示例演示如何设置、获取和删除自定义标头。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-124">The following example shows how to set, get, and remove custom headers.</span></span>

```js
// Set custom internet headers.
function setCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "x-preferred-fruit": "orange", "x-preferred-vegetable": "broccoli", "x-best-vegetable": "spinach" },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.getAsync(
    ["x-preferred-fruit", "x-preferred-vegetable", "x-best-vegetable", "x-nonexistent-header"],
    getCallback
  );
}

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected headers: " + JSON.stringify(asyncResult.value));
  } else {
    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
  }
}

// Remove custom internet headers.
function removeSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.removeAsync(
    ["x-best-vegetable", "x-nonexistent-header"],
    removeCallback);
}

function removeCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully removed selected headers");
  } else {
    console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
  }
}

setCustomHeaders();
getSelectedCustomHeaders();
removeSelectedCustomHeaders();
getSelectedCustomHeaders();

/* Sample output:
Successfully set headers
Selected headers: {"x-best-vegetable":"spinach","x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
Successfully removed selected headers
Selected headers: {"x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
*/
```

## <a name="get-internet-headers-while-reading-a-message"></a><span data-ttu-id="ef1dc-125">在阅读邮件时获取 internet 邮件头</span><span class="sxs-lookup"><span data-stu-id="ef1dc-125">Get internet headers while reading a message</span></span>

<span data-ttu-id="ef1dc-126">尝试调用[getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-)以在阅读模式下获取当前邮件的 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-126">Try calling [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) to get internet headers on the current message in Read mode.</span></span>

### <a name="get-sender-preferences-from-current-mime-headers-example"></a><span data-ttu-id="ef1dc-127">从当前 MIME 标头获取发件人首选项示例</span><span class="sxs-lookup"><span data-stu-id="ef1dc-127">Get sender preferences from current MIME headers example</span></span>

<span data-ttu-id="ef1dc-128">根据上一节中的示例，以下代码演示如何从当前电子邮件的 MIME 标头中获取发件人的首选项。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-128">Building on the example from the previous section, the following code shows how to get the sender's preferences from the current email's MIME headers.</span></span>

```js
Office.context.mailbox.item.getAllInternetHeadersAsync(getCallback);

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Sender's preferred fruit: " + asyncResult.value.match(/x-preferred-fruit:.*/gim)[0].slice(19));
    console.log("Sender's preferred vegetable: " + asyncResult.value.match(/x-preferred-vegetable:.*/gim)[0].slice(23));
  } else {
    console.log("Error getting preferences from header: " + JSON.stringify(asyncResult.error));
  }
}

/* Sample output:
Sender's preferred fruit: orange
Sender's preferred vegetable: broccoli
*/
```

> [!IMPORTANT]
> <span data-ttu-id="ef1dc-129">此示例适用于简单的情况。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-129">This sample works for simple cases.</span></span> <span data-ttu-id="ef1dc-130">若要获取更复杂的信息检索（如[RFC 2822](https://tools.ietf.org/html/rfc2822)中所述的多实例标头或折叠的值），请尝试使用相应的 MIME 分析库。</span><span class="sxs-lookup"><span data-stu-id="ef1dc-130">For more complex information retrieval (e.g., multi-instance headers or folded values as described in [RFC 2822](https://tools.ietf.org/html/rfc2822)), try using an appropriate MIME-parsing library.</span></span>

## <a name="see-also"></a><span data-ttu-id="ef1dc-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ef1dc-131">See also</span></span>

- [<span data-ttu-id="ef1dc-132">获取和设置 Outlook 加载项的元数据</span><span class="sxs-lookup"><span data-stu-id="ef1dc-132">Get and set add-in metadata for an Outlook add-in</span></span>](metadata-for-an-outlook-add-in.md)
