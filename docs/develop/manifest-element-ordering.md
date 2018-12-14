---
title: 如何查找清单元素的正确顺序
description: 了解如何查找在父元素中放置子元素的正确顺序。
ms.date: 11/16/2018
ms.openlocfilehash: 3efc95926b7562b0e68bbb6f4b13c47cc4ae6824
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270612"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a><span data-ttu-id="b9184-103">如何查找清单元素的正确顺序</span><span class="sxs-lookup"><span data-stu-id="b9184-103">How to find the proper order of manifest elements</span></span>

<span data-ttu-id="b9184-104">Office 外接程序清单中的 XML 元素必须位于正确父元素下，*且*在父元素下以特定的相对顺序存在。</span><span class="sxs-lookup"><span data-stu-id="b9184-104">The XML elements in the manifest of an Office Add-in must be under the proper parent element *and* in a specific order, relative to each other, under the parent.</span></span>

<span data-ttu-id="b9184-105">所需的排序在 [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 文件夹的 XSD 文件中指定。</span><span class="sxs-lookup"><span data-stu-id="b9184-105">The required ordering is specified in the XSD files in the [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) folder.</span></span> <span data-ttu-id="b9184-106">XSD 文件分类存放在对应任务窗格、内容和邮件三类外接程序的子文件夹中。</span><span class="sxs-lookup"><span data-stu-id="b9184-106">The XSD files are categorized into subfolders for taskpane, content, and mail add-ins.</span></span>

<span data-ttu-id="b9184-107">例如，在 `<OfficeApp>` 元素中，`<Id>`、`<Version>`、`<ProviderName>` 必须按此顺序出现。</span><span class="sxs-lookup"><span data-stu-id="b9184-107">For example, In the `<OfficeApp>` element, the `<Id>`, `<Version>`, `<ProviderName>` must appear in that order.</span></span> <span data-ttu-id="b9184-108">如果添加了 `<AlternateId>` 元素，则其必须位于 `<Id>` 和 `<Version>` 元素之间。</span><span class="sxs-lookup"><span data-stu-id="b9184-108">If an `<AlternateId>` element is added, it must be between the `<Id>` and `<Version>` element.</span></span> <span data-ttu-id="b9184-109">如果任何元素的顺序出错，清单将无效并且你的外接程序将无法加载。</span><span class="sxs-lookup"><span data-stu-id="b9184-109">Your manifest will not be valid and your add-in will not load, if any element is in the wrong order.</span></span>

> [!NOTE]
> <span data-ttu-id="b9184-110">当元素顺序被打乱时，[Office 外接程序验证程序](/office/dev/add-ins/testing/troubleshoot-manifest#validate-your-manifest-with-the-office-add-in-validator)将使用与元素位于错误父级下时相同的错误消息。</span><span class="sxs-lookup"><span data-stu-id="b9184-110">The [Office Add-in Validator](/office/dev/add-ins/testing/troubleshoot-manifest#validate-your-manifest-with-the-office-add-in-validator) uses the same error message when an element is out-of-order as it does when an element is under the wrong parent.</span></span> <span data-ttu-id="b9184-111">该错误消息会提示子元素不是父元素的有效子级。</span><span class="sxs-lookup"><span data-stu-id="b9184-111">The error says the child element is is not a valid child of the parent element.</span></span> <span data-ttu-id="b9184-112">如果出现此类错误，而子元素的参考文档却指示它对父级*是*有效的，则问题很可能是子级的放置顺序出现了错误。</span><span class="sxs-lookup"><span data-stu-id="b9184-112">If you get such an error but the reference documentation for the child element indicates that it *is* valid for the parent, then the problem is likely that the child has been placed in the wrong order.</span></span>

<span data-ttu-id="b9184-113">若要查找给定父元素的子元素的正确顺序，请执行以下步骤。</span><span class="sxs-lookup"><span data-stu-id="b9184-113">To find the correct order for the child elements of a given parent element, take the following steps.</span></span> <span data-ttu-id="b9184-114">（这是一个简化的过程，因为 XSD 文件非常复杂。</span><span class="sxs-lookup"><span data-stu-id="b9184-114">(This is a simplified process, as XSD files are quite complex.</span></span> <span data-ttu-id="b9184-115">完全解析 XSD 文件不在本文的讨论范围之列。）</span><span class="sxs-lookup"><span data-stu-id="b9184-115">Fully parsing XSD files is out of the scope of this document.)</span></span>

1. <span data-ttu-id="b9184-116">打开 [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) 下的子文件夹，以获取你正在创建的外接程序的类型。</span><span class="sxs-lookup"><span data-stu-id="b9184-116">Open the subfolder under [Schemas](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) for the type of add-in that you are creating.</span></span> 
2. <span data-ttu-id="b9184-117">打开 XSD 文件，其中父元素被定义为复杂类型。</span><span class="sxs-lookup"><span data-stu-id="b9184-117">Open the XSD file where the parent element is defined as a complex type.</span></span> <span data-ttu-id="b9184-118">如果你不知道哪个文件具有该定义，则可能必须对多个文件执行步骤 3，直到找到它为止。</span><span class="sxs-lookup"><span data-stu-id="b9184-118">If you don't know which file has the definition, you may have to do step 3 on multiple files until you find it.</span></span>
3. <span data-ttu-id="b9184-119">搜索 `<xs:complexType name="PARENT_ELEMENT">`，其中 PARENT_ELEMENT 是该父元素的名称。</span><span class="sxs-lookup"><span data-stu-id="b9184-119">Search for `<xs:complexType name="PARENT_ELEMENT">`, where PARENT_ELEMENT is the name of the parent element.</span></span>
4. <span data-ttu-id="b9184-120">在 PARENT_ELEMENT 的定义中，（通常）有一个名为 `<xs:sequence>` 的元素。</span><span class="sxs-lookup"><span data-stu-id="b9184-120">Inside the definition for the PARENT_ELEMENT, there is (usually) an element called `<xs:sequence>`.</span></span> <span data-ttu-id="b9184-121">以下是 [TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd) 中对 `<SuperTip>` 的定义。</span><span class="sxs-lookup"><span data-stu-id="b9184-121">The following is the definition for `<SuperTip>` from [TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd).</span></span>

```xml
  <xs:complexType name="Supertip">
    <xs:annotation>
      <xs:documentation>
        Specifies the super tip for this control.
      </xs:documentation>
    </xs:annotation>
    <xs:sequence>
      <xs:element name="Title" type="bt:ShortResourceReference" minOccurs="1" maxOccurs="1" />
      <xs:element name="Description" type="bt:LongResourceReference" minOccurs="1" maxOccurs="1" />
    </xs:sequence>
  </xs:complexType>
```

<span data-ttu-id="b9184-122">`<xs:sequence>` *按照子元素必须出现的顺序*列出了可能的子元素。</span><span class="sxs-lookup"><span data-stu-id="b9184-122">The `<xs:sequence>` lists the possible child elements, *in the order in which they must appear*.</span></span> <span data-ttu-id="b9184-123">但这并*不*意味着它们全都是必需的。</span><span class="sxs-lookup"><span data-stu-id="b9184-123">This does *not* mean all of them are mandatory.</span></span> <span data-ttu-id="b9184-124">如果某个子元素的 `minOccurs` 值为 **0**，则该子元素是可选的。</span><span class="sxs-lookup"><span data-stu-id="b9184-124">If the `minOccurs` value for a child element is **0**, then the child element is optional.</span></span> <span data-ttu-id="b9184-125">*但如果该元素存在，则必须以由 `<xs:sequence>` 元素指定的顺序出现*。</span><span class="sxs-lookup"><span data-stu-id="b9184-125">*But if it is present, it must be in the order specified by the `<xs:sequence>` element*.</span></span>

<span data-ttu-id="b9184-126">如果没有 `<xs:sequence>` 元素，或者*有*该子元素但未列出（即使子元素的参考文档指示它对父级*是*有效的）；则在 XSD 文件中的其他位置通过其他子元素对父元素的复杂类型定义进行了扩展。</span><span class="sxs-lookup"><span data-stu-id="b9184-126">If there is no `<xs:sequence>` element, or there *is* but the child element is not listed (even though the reference documentation for the child element indicates that it *is* valid for the parent); then the parent element's complex type definition has been extended with additional child elements somewhere else in the XSD file.</span></span> <span data-ttu-id="b9184-127">例如，`OfficeApp` 复杂类型的定义未将 `Requirements` 列为可能的子级。</span><span class="sxs-lookup"><span data-stu-id="b9184-127">For example, the definition for the `OfficeApp` complex type does not list `Requirements` as a possible child.</span></span> <span data-ttu-id="b9184-128">但在文件的稍后部分（在 `TaskPaneApp` 复杂类型的定义中），对 `OfficeApp` 的定义进行了扩展，并添加了 `Requirements` 作为其他有效子级。</span><span class="sxs-lookup"><span data-stu-id="b9184-128">But later in the file (within the definition for the `TaskPaneApp` complex type), the definition of `OfficeApp` is extended and `Requirements` is added as an additional valid child.</span></span>

<span data-ttu-id="b9184-129">若要查找扩展的定义，请按照以下步骤操作：</span><span class="sxs-lookup"><span data-stu-id="b9184-129">To find the extended definitions follow these steps:</span></span>

1. <span data-ttu-id="b9184-130">从文件的顶部开始，搜索 `<xs:extension base="PARENT_ELEMENT">`，其中 PARENT_ELEMENT 是父元素的名称。</span><span class="sxs-lookup"><span data-stu-id="b9184-130">Starting at the top of the file, search for `<xs:extension base="PARENT_ELEMENT">`, where PARENT_ELEMENT is the name of the parent element.</span></span> <span data-ttu-id="b9184-131">可能存在多个扩展。</span><span class="sxs-lookup"><span data-stu-id="b9184-131">There may be more than one extension.</span></span>
2. <span data-ttu-id="b9184-132">查找与你正在使用的上下文相关的扩展。</span><span class="sxs-lookup"><span data-stu-id="b9184-132">Find the extension that is relevant to the context in which you are working.</span></span> <span data-ttu-id="b9184-133">例如，`OfficeApp` 复杂类型在 `ContentApp` 和 `MailApp` 复杂类型内进行了扩展，同时也在 `TaskPaneApp` 复杂类型内进行了扩展。</span><span class="sxs-lookup"><span data-stu-id="b9184-133">For example, the `OfficeApp` complex type is extended within the `ContentApp` and `MailApp` complex types as well as within the `TaskPaneApp` complex type.</span></span>

<span data-ttu-id="b9184-134">文件中的每个 `<xs:extension base="PARENT_ELEMENT">` 都有自己的 `<xs:sequence>`，后者会列出父级的其他有效子元素。</span><span class="sxs-lookup"><span data-stu-id="b9184-134">Each `<xs:extension base="PARENT_ELEMENT">` in the file has its own `<xs:sequence>` that lists additional valid child elements for the parent.</span></span> <span data-ttu-id="b9184-135">扩展列表上的子元素必须始终位于父级复杂类型定义中原始列表的子元素*之后*。</span><span class="sxs-lookup"><span data-stu-id="b9184-135">Child elements on an extended list must always be *after* the child elements in the original list in the parent's complex type definition.</span></span>

## <a name="see-also"></a><span data-ttu-id="b9184-136">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b9184-136">See also</span></span>

- [<span data-ttu-id="b9184-137">Office 外接程序清单的架构参考 (v1.1)</span><span class="sxs-lookup"><span data-stu-id="b9184-137">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
