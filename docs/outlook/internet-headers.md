---
title: 获取和设置 internet 标头
description: 如何：在 Outlook 外接程序中获取和设置邮件的 internet 邮件头。
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: 1b6bdbbe77998ce92ea1b1b43874a32a30aa160a
ms.sourcegitcommit: 0fdb78cefa669b727b817614a4147a46d249a0ed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/28/2020
ms.locfileid: "43930286"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a><span data-ttu-id="619bb-103">在 Outlook 外接程序中获取和设置邮件的 internet 邮件头</span><span class="sxs-lookup"><span data-stu-id="619bb-103">Get and set internet headers on a message in an Outlook add-in</span></span>

## <a name="background"></a><span data-ttu-id="619bb-104">背景</span><span class="sxs-lookup"><span data-stu-id="619bb-104">Background</span></span>

<span data-ttu-id="619bb-105">Outlook 外接程序开发中的一个常见要求是，将与外接程序关联的自定义属性存储在不同的级别。</span><span class="sxs-lookup"><span data-stu-id="619bb-105">A common requirement in Outlook add-ins development is to store custom properties associated with an add-in at different levels.</span></span> <span data-ttu-id="619bb-106">目前，自定义属性存储在项目或邮箱级别。</span><span class="sxs-lookup"><span data-stu-id="619bb-106">At present, custom properties are stored at the item or mailbox level.</span></span>

- <span data-ttu-id="619bb-107">项目级别-适用于特定项目的属性，使用[CustomProperties](/javascript/api/outlook/office.customproperties)对象。</span><span class="sxs-lookup"><span data-stu-id="619bb-107">Item level - For properties that apply to a specific item, use the [CustomProperties](/javascript/api/outlook/office.customproperties) object.</span></span> <span data-ttu-id="619bb-108">例如，存储与发送电子邮件的人员关联的客户代码。</span><span class="sxs-lookup"><span data-stu-id="619bb-108">For example, store a customer code associated with the person who sent the email.</span></span>
- <span data-ttu-id="619bb-109">邮箱级别-对于应用于用户邮箱中的所有邮件项目的属性，使用[RoamingSettings](/javascript/api/outlook/office.roamingsettings)对象。</span><span class="sxs-lookup"><span data-stu-id="619bb-109">Mailbox level - For properties that apply to all the mail items in the user's mailbox, use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object.</span></span> <span data-ttu-id="619bb-110">例如，存储用户首选项以按特定比例显示温度。</span><span class="sxs-lookup"><span data-stu-id="619bb-110">For example, store a user's preference to show the temperature in a particular scale.</span></span>

<span data-ttu-id="619bb-111">在项目离开 Exchange 服务器后，这两种类型的属性都不会保留，因此电子邮件收件人无法获取项目上设置的任何属性。</span><span class="sxs-lookup"><span data-stu-id="619bb-111">Both types of properties are not preserved after the item leaves the Exchange server so the email recipients can't get any properties set on the item.</span></span> <span data-ttu-id="619bb-112">因此，开发人员无法访问这些设置或其他 MIME 属性以实现更好的阅读方案。</span><span class="sxs-lookup"><span data-stu-id="619bb-112">Therefore, developers can't access those settings or other MIME properties to enable better read scenarios.</span></span>

<span data-ttu-id="619bb-113">虽然有一种方法可以将 internet 标头设置为 EWS 请求，但在某些情况下，不能进行 EWS 请求。</span><span class="sxs-lookup"><span data-stu-id="619bb-113">While there's a way for you to set the internet headers through EWS requests, in some scenarios making an EWS request won't work.</span></span> <span data-ttu-id="619bb-114">例如，在 Outlook 桌面的撰写模式下，项目 id 在缓存模式下 `saveAsync` 不会同步。</span><span class="sxs-lookup"><span data-stu-id="619bb-114">For example, in Compose mode on Outlook desktop, the item id isn't synced on `saveAsync` in cached mode.</span></span>

> [!TIP]
> <span data-ttu-id="619bb-115">请参阅[获取和设置 Outlook 外接程序的外接程序元数据](metadata-for-an-outlook-add-in.md)，以了解有关使用这些选项的详细信息。</span><span class="sxs-lookup"><span data-stu-id="619bb-115">See [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md) to learn more about using these options.</span></span>

## <a name="purpose-of-the-internet-headers-api"></a><span data-ttu-id="619bb-116">Internet 标头 API 的用途</span><span class="sxs-lookup"><span data-stu-id="619bb-116">Purpose of the internet headers API</span></span>

<span data-ttu-id="619bb-117">在[要求集 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)中引入，internet 标头 api 使开发人员能够：</span><span class="sxs-lookup"><span data-stu-id="619bb-117">Introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), the internet headers APIs enable developers to:</span></span>

- <span data-ttu-id="619bb-118">戳在所有客户端上保留 Exchange 后保留的电子邮件的信息。</span><span class="sxs-lookup"><span data-stu-id="619bb-118">Stamp information on an email that persists after it leaves Exchange across all clients.</span></span>
- <span data-ttu-id="619bb-119">阅读有关在邮件读取应用场景中的所有客户端上的电子邮件保留后保留的电子邮件的信息。</span><span class="sxs-lookup"><span data-stu-id="619bb-119">Read information on an email that persisted after the email left Exchange across all clients in mail read scenarios.</span></span>
- <span data-ttu-id="619bb-120">访问电子邮件的整个 MIME 标头。</span><span class="sxs-lookup"><span data-stu-id="619bb-120">Access the entire MIME header of the email.</span></span>

![Internet 标头的图示。](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a><span data-ttu-id="619bb-126">在撰写邮件时设置 internet 邮件头</span><span class="sxs-lookup"><span data-stu-id="619bb-126">Set internet headers while composing a message</span></span>

<span data-ttu-id="619bb-127">尝试使用[internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders)属性来管理在撰写模式下放置在当前邮件上的自定义 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="619bb-127">Try using the [item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) property to manage the custom internet headers you place on the current message in Compose mode.</span></span>

### <a name="set-get-and-remove-custom-headers-example"></a><span data-ttu-id="619bb-128">设置、获取和删除自定义标头示例</span><span class="sxs-lookup"><span data-stu-id="619bb-128">Set, get, and remove custom headers example</span></span>

<span data-ttu-id="619bb-129">下面的示例演示如何设置、获取和删除自定义标头。</span><span class="sxs-lookup"><span data-stu-id="619bb-129">The following example shows how to set, get, and remove custom headers.</span></span>

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

## <a name="get-internet-headers-while-reading-a-message"></a><span data-ttu-id="619bb-130">在阅读邮件时获取 internet 邮件头</span><span class="sxs-lookup"><span data-stu-id="619bb-130">Get internet headers while reading a message</span></span>

<span data-ttu-id="619bb-131">尝试调用[getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-)以在阅读模式下获取当前邮件的 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="619bb-131">Try calling [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) to get internet headers on the current message in Read mode.</span></span>

### <a name="get-sender-preferences-from-current-mime-headers-example"></a><span data-ttu-id="619bb-132">从当前 MIME 标头获取发件人首选项示例</span><span class="sxs-lookup"><span data-stu-id="619bb-132">Get sender preferences from current MIME headers example</span></span>

<span data-ttu-id="619bb-133">根据上一节中的示例，以下代码演示如何从当前电子邮件的 MIME 标头中获取发件人的首选项。</span><span class="sxs-lookup"><span data-stu-id="619bb-133">Building on the example from the previous section, the following code shows how to get the sender's preferences from the current email's MIME headers.</span></span>

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
> <span data-ttu-id="619bb-134">此示例适用于简单的情况。</span><span class="sxs-lookup"><span data-stu-id="619bb-134">This sample works for simple cases.</span></span> <span data-ttu-id="619bb-135">若要获取更复杂的信息检索（例如，多实例标头或折叠的值（如[RFC 2822](https://tools.ietf.org/html/rfc2822)中所述），请尝试使用相应的 MIME 分析库。</span><span class="sxs-lookup"><span data-stu-id="619bb-135">For more complex information retrieval (for example, multi-instance headers or folded values as described in [RFC 2822](https://tools.ietf.org/html/rfc2822)), try using an appropriate MIME-parsing library.</span></span>

## <a name="recommended-practices"></a><span data-ttu-id="619bb-136">建议的做法</span><span class="sxs-lookup"><span data-stu-id="619bb-136">Recommended practices</span></span>

<span data-ttu-id="619bb-137">目前，internet 邮件头是用户邮箱的有限资源。</span><span class="sxs-lookup"><span data-stu-id="619bb-137">Currently, internet headers are a finite resource on a user's mailbox.</span></span> <span data-ttu-id="619bb-138">当配额耗尽时，您不能在该邮箱上创建更多的 internet 标头，这可能导致依赖于此的客户端的意外行为能够正常运行。</span><span class="sxs-lookup"><span data-stu-id="619bb-138">When the quota is exhausted, you can't create any more internet headers on that mailbox, which can result in unexpected behavior from clients that rely on this to function.</span></span>

<span data-ttu-id="619bb-139">在外接程序中创建 internet 邮件头时，请应用以下准则。</span><span class="sxs-lookup"><span data-stu-id="619bb-139">Apply the following guidelines when you create internet headers in your add-in.</span></span>

- <span data-ttu-id="619bb-140">创建所需的最小标头数。</span><span class="sxs-lookup"><span data-stu-id="619bb-140">Create the minimum number of headers required.</span></span>
- <span data-ttu-id="619bb-141">名称标头，以便以后可以重复使用和更新其值。</span><span class="sxs-lookup"><span data-stu-id="619bb-141">Name headers so that you can reuse and update their values later.</span></span> <span data-ttu-id="619bb-142">因此，应避免以变量方式命名标头（例如，基于用户输入、时间戳等）。</span><span class="sxs-lookup"><span data-stu-id="619bb-142">As such, avoid naming headers in a variable manner (for example, based on user input, timestamp, etc.).</span></span>

## <a name="see-also"></a><span data-ttu-id="619bb-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="619bb-143">See also</span></span>

- [<span data-ttu-id="619bb-144">获取和设置 Outlook 加载项的元数据</span><span class="sxs-lookup"><span data-stu-id="619bb-144">Get and set add-in metadata for an Outlook add-in</span></span>](metadata-for-an-outlook-add-in.md)
