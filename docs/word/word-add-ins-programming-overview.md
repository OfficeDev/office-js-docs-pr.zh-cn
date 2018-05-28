---
title: Word ?????
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 63605c18f7e1b3eae2c542aef236372819bc2e6f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="a842c-102">Word ?????</span><span class="sxs-lookup"><span data-stu-id="a842c-102">Word add-ins overview</span></span>

<span data-ttu-id="a842c-p101">?????????? Word ?????????????????????????????????? Word ??????????????? Office ?????????? Word JavaScript API ???? Office ? JavaScript API??????? Windows ?????Mac ?????? Word ????</span><span class="sxs-lookup"><span data-stu-id="a842c-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the JavaScript API for Office, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="a842c-p102">Word ????? [Office ??????](../overview/office-add-ins.md)??????????????????????? Word ????????? JavaScript ?? Word ?????????????????????????????? Word ????????? Word ??????????????????? Word ?????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span> 

> [!NOTE]
> <span data-ttu-id="a842c-p103">???????????????[??](../publish/publish.md)? AppSource?????? [AppSource ????](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)??????????????????????????????????????????[? 4.12 ??](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)?? [Office ?????????](../overview/office-add-in-availability.md)????</span><span class="sxs-lookup"><span data-stu-id="a842c-p103">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="a842c-113">?????????????????? Word ????</span><span class="sxs-lookup"><span data-stu-id="a842c-113">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="a842c-114">*? 1?? Word ????????????*</span><span class="sxs-lookup"><span data-stu-id="a842c-114">*Figure 1. Add-in running in a task pane in Word*</span></span>

![? Word ?????????????](../images/word-add-in-show-host-client.png)

<span data-ttu-id="a842c-p104">Word ???? (1) ???????? Word ?? (2) ???? JavaScript ????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-p104">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

<span data-ttu-id="a842c-p105">???? Word ?????????? Web ??????? ASP.NET?NodeJS ? Python???????????????Ember?Backbone?Angular?React????????? VanillaJS ??????????? Azure ???[??](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md)??????</span><span class="sxs-lookup"><span data-stu-id="a842c-p105">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md) and host your application.</span></span>

<span data-ttu-id="a842c-p106">?? Word JavaScript API ???????? Word ????????????? API ???????????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-p106">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="a842c-122">??? Windows ? Word 2013</span><span class="sxs-lookup"><span data-stu-id="a842c-122">Word 2013 for Windows</span></span>
* <span data-ttu-id="a842c-123">??? Windows ? Word 2016</span><span class="sxs-lookup"><span data-stu-id="a842c-123">Word 2016 for Windows</span></span>
* <span data-ttu-id="a842c-124">Word Online</span><span class="sxs-lookup"><span data-stu-id="a842c-124">Word Online</span></span>
* <span data-ttu-id="a842c-125">??? Mac ? Word 2016</span><span class="sxs-lookup"><span data-stu-id="a842c-125">Word 2016 for Mac</span></span>
* <span data-ttu-id="a842c-126">??? iOS ? Word</span><span class="sxs-lookup"><span data-stu-id="a842c-126">Word for iOS</span></span>

<span data-ttu-id="a842c-p107">??????????????????????? Word ?????????????? [Office ????????????](../overview/office-add-in-availability.md)?</span><span class="sxs-lookup"><span data-stu-id="a842c-p107">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="a842c-129">??? Word ? JavaScript API</span><span class="sxs-lookup"><span data-stu-id="a842c-129">JavaScript APIs for Word</span></span>

<span data-ttu-id="a842c-p108">??? JavaScript Api ???? Word ???????????????????[??? Office ? JavaScript API](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word)?? Office 2013 ???????????? API --????????????? Office ??????????????? API ???????</span><span class="sxs-lookup"><span data-stu-id="a842c-p108">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document. The first is the [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word), which was introduced in Office 2013. This is a shared API -- many of the objects can be used in add-ins hosted by two or more Office clients. This API uses callbacks extensively.</span></span> 

<span data-ttu-id="a842c-p109">???? [Word JavaScript API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)????????????????????? Mac ? Windows ? Word 2016 ? Word ???????????????????????? Word ????[??](https://dev.office.com/reference/add-ins/word/body)?[????](https://dev.office.com/reference/add-ins/word/contentcontrol)?[????](https://dev.office.com/reference/add-ins/word/inlinepicture)?[??](https://dev.office.com/reference/add-ins/word/paragraph)???????Word JavaScript API ?? TypeScript ??? vsdoc ?????????? IDE ????????</span><span class="sxs-lookup"><span data-stu-id="a842c-p109">The second is the [Word JavaScript API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 for Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](https://dev.office.com/reference/add-ins/word/body), [content controls](https://dev.office.com/reference/add-ins/word/contentcontrol), [inline pictures](https://dev.office.com/reference/add-ins/word/inlinepicture), and [paragraphs](https://dev.office.com/reference/add-ins/word/paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="a842c-p110">????? Word ???????????? Office ? JavaScript API????????? Word JavaScript API??????????????????? [API ????](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word)?</span><span class="sxs-lookup"><span data-stu-id="a842c-p110">Currently, all Word clients support the shared JavaScript API for Office, and most clients support the Word JavaScript API. For details about supported clients, see the [API reference documentation](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word).</span></span>

<span data-ttu-id="a842c-p111">????? Word JavaScript API ????????????????????????????? Word JavaScript API?</span><span class="sxs-lookup"><span data-stu-id="a842c-p111">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="a842c-142">?? Word ???????</span><span class="sxs-lookup"><span data-stu-id="a842c-142">Access the objects in a Word document.</span></span>

<span data-ttu-id="a842c-143">??????????????????? Office ? JavaScript API?</span><span class="sxs-lookup"><span data-stu-id="a842c-143">Use the shared JavaScript API for Office when you need to:</span></span>

* <span data-ttu-id="a842c-144">?? Word 2013?</span><span class="sxs-lookup"><span data-stu-id="a842c-144">Target Word 2013.</span></span>
* <span data-ttu-id="a842c-145">????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-145">Perform initial actions for the application.</span></span>
* <span data-ttu-id="a842c-146">?????????</span><span class="sxs-lookup"><span data-stu-id="a842c-146">Check the supported requirement set.</span></span>
* <span data-ttu-id="a842c-147">?????????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-147">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="a842c-148">???????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-148">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="a842c-149">????? XML ???</span><span class="sxs-lookup"><span data-stu-id="a842c-149">Use custom XML parts.</span></span>
* <span data-ttu-id="a842c-150">????????</span><span class="sxs-lookup"><span data-stu-id="a842c-150">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a842c-151">????</span><span class="sxs-lookup"><span data-stu-id="a842c-151">Next steps</span></span>

<span data-ttu-id="a842c-p112">??????? Word ?????????[???? Word ???](word-add-ins.md)???????????[????](http://dev.office.com/getting-started/addins?product=Word)????[?????](../develop/add-in-manifests.md)???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-p112">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). You can also try our interactive [Get started experience](http://dev.office.com/getting-started/addins?product=Word). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="a842c-156">????????????? Word ??????????????????????[????](../design/add-in-design.md)?[????](../concepts/add-in-development-best-practices.md)?</span><span class="sxs-lookup"><span data-stu-id="a842c-156">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="a842c-157">???????????[??](../publish/publish.md)??????????? AppSource?</span><span class="sxs-lookup"><span data-stu-id="a842c-157">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="whats-coming-up-for-word-add-ins"></a><span data-ttu-id="a842c-158">Word ????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-158">What's coming up for Word add-ins?</span></span>

<span data-ttu-id="a842c-p113">???????????? Word ????? API????????? API?????? [API ?????](https://dev.office.com/reference/add-ins/openspec)???????????????? Word JavaScript API ??????????????????</span><span class="sxs-lookup"><span data-stu-id="a842c-p113">As we design and develop new APIs for Word add-ins, we'll make them available for your feedback on our [API open specifications](https://dev.office.com/reference/add-ins/openspec) page. Find out what new features are in the pipeline for the Word JavaScript APIs, and provide your input on our design specifications.</span></span>

## <a name="see-also"></a><span data-ttu-id="a842c-161">????</span><span class="sxs-lookup"><span data-stu-id="a842c-161">See also</span></span>

* [<span data-ttu-id="a842c-162">Office ???????</span><span class="sxs-lookup"><span data-stu-id="a842c-162">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="a842c-163">Word JavaScript API ??</span><span class="sxs-lookup"><span data-stu-id="a842c-163">Word JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)

