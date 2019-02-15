---
title: OneNote JavaScript API 概述
description: ''
ms.date: 10/09/2018
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: f8fed0104412f60ec59146ef7820be958047d1f3
ms.sourcegitcommit: 70ef38a290c18a1d1a380fd02b263470207a5dc6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/15/2019
ms.locfileid: "30052740"
---
# <a name="onenote-javascript-api-overview"></a><span data-ttu-id="27d8c-102">OneNote JavaScript API 概述</span><span class="sxs-lookup"><span data-stu-id="27d8c-102">OneNote JavaScript API overview</span></span>

<span data-ttu-id="27d8c-103">适用于：OneNote Online</span><span class="sxs-lookup"><span data-stu-id="27d8c-103">Applies to: OneNote Online</span></span>

<span data-ttu-id="27d8c-104">下面的链接展示了 API 中的高级 OneNote 对象。</span><span class="sxs-lookup"><span data-stu-id="27d8c-104">The following links show the high level OneNote objects available in the API.</span></span> <span data-ttu-id="27d8c-105">每个对象页面链接包含对象可用的属性、事件和方法的描述。</span><span class="sxs-lookup"><span data-stu-id="27d8c-105">Each object page link contains a description of the properties, events, and methods available on the object.</span></span> <span data-ttu-id="27d8c-106">如需了解详细信息，请浏览相应链接。</span><span class="sxs-lookup"><span data-stu-id="27d8c-106">Explore these links to learn more.</span></span> 
    
- <span data-ttu-id="27d8c-107">[Application](/javascript/api/onenote/onenote.application)：用于访问所有全局可寻址的 OneNote 对象（如活动笔记本和活动分区）的顶级对象。</span><span class="sxs-lookup"><span data-stu-id="27d8c-107">[Application](/javascript/api/onenote/onenote.application): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.</span></span>

- <span data-ttu-id="27d8c-p102">[笔记本](/javascript/api/onenote/onenote.notebook)：一个笔记本。笔记本包含分区组合和分区。</span><span class="sxs-lookup"><span data-stu-id="27d8c-p102">[Notebook](/javascript/api/onenote/onenote.notebook): A notebook. Notebooks contain section groups and sections.</span></span>
    - <span data-ttu-id="27d8c-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection)：笔记本的集合。</span><span class="sxs-lookup"><span data-stu-id="27d8c-110">[NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): A collection of notebooks.</span></span>

- <span data-ttu-id="27d8c-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup)：一个分区组。分区组包含分区组和分区。</span><span class="sxs-lookup"><span data-stu-id="27d8c-p103">[SectionGroup](/javascript/api/onenote/onenote.sectiongroup): A section group. Section groups contain section groups and sections.</span></span>
    - <span data-ttu-id="27d8c-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection)：分区组的集合。</span><span class="sxs-lookup"><span data-stu-id="27d8c-113">[SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): A collection of section groups.</span></span>

- <span data-ttu-id="27d8c-p104">[Section](/javascript/api/onenote/onenote.section)：一个分区。分区包含页面。</span><span class="sxs-lookup"><span data-stu-id="27d8c-p104">[Section](/javascript/api/onenote/onenote.section): A section. Sections contain pages.</span></span>
    - <span data-ttu-id="27d8c-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection)：分区的集合。</span><span class="sxs-lookup"><span data-stu-id="27d8c-116">[SectionCollection](/javascript/api/onenote/onenote.sectioncollection): A collection of sections.</span></span>

- <span data-ttu-id="27d8c-p105">[Page](/javascript/api/onenote/onenote.page)：一个页面。页面包含 PageContent 对象。</span><span class="sxs-lookup"><span data-stu-id="27d8c-p105">[Page](/javascript/api/onenote/onenote.page): A page. Pages contain PageContent objects.</span></span>
    - <span data-ttu-id="27d8c-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection)：页面的集合。</span><span class="sxs-lookup"><span data-stu-id="27d8c-119">[PageCollection](/javascript/api/onenote/onenote.pagecollection): A collection of pages.</span></span>

- <span data-ttu-id="27d8c-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent)：页面上包含内容类型的顶级地区，例如 Outline 或 Image。可在页面上为 PageContent 对象分配一个位置。</span><span class="sxs-lookup"><span data-stu-id="27d8c-p106">[PageContent](/javascript/api/onenote/onenote.pagecontent): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.</span></span>
    - <span data-ttu-id="27d8c-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection)：PageContent 对象的集合，表示页面的内容。</span><span class="sxs-lookup"><span data-stu-id="27d8c-122">[PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): A collection of PageContent objects, which represents the contents of a page.</span></span>

- <span data-ttu-id="27d8c-p107">[Outline](/javascript/api/onenote/onenote.outline)：Paragraph 对象的容器。Outline 是 PageContent 对象的直接子级。</span><span class="sxs-lookup"><span data-stu-id="27d8c-p107">[Outline](/javascript/api/onenote/onenote.outline): A container for Paragraph objects. An Outline is a direct child of a PageContent object.</span></span>

- <span data-ttu-id="27d8c-p108">[Image](/javascript/api/onenote/onenote.image)：Image 对象。Image 可以是 PageContent 对象或 Paragraph 的直接子级。</span><span class="sxs-lookup"><span data-stu-id="27d8c-p108">[Image](/javascript/api/onenote/onenote.image): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.</span></span>

- <span data-ttu-id="27d8c-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph)：页面上可见内容的容器。Paragraph 是 Outline 的直接子级。</span><span class="sxs-lookup"><span data-stu-id="27d8c-p109">[Paragraph](/javascript/api/onenote/onenote.paragraph): A container for the visible content on a page. A Paragraph is a direct child of an Outline.</span></span>
    - <span data-ttu-id="27d8c-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection)：Outline 中 Paragraph 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="27d8c-129">[ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): A collection of Paragraph objects in an Outline.</span></span>

- <span data-ttu-id="27d8c-130">[RichText](/javascript/api/onenote/onenote.richtext)：RichText 对象。</span><span class="sxs-lookup"><span data-stu-id="27d8c-130">[RichText](/javascript/api/onenote/onenote.richtext): A RichText object.</span></span>

- <span data-ttu-id="27d8c-131">[表格](/javascript/api/onenote/onenote.table)：TableRow 对象的容器。</span><span class="sxs-lookup"><span data-stu-id="27d8c-131">[Table](/javascript/api/onenote/onenote.table): A container for TableRow objects.</span></span>

- <span data-ttu-id="27d8c-132">[TableRow](/javascript/api/onenote/onenote.tablerow)：TableCell 对象的容器。</span><span class="sxs-lookup"><span data-stu-id="27d8c-132">[TableRow](/javascript/api/onenote/onenote.tablerow): A container for TableCell objects.</span></span>
    - <span data-ttu-id="27d8c-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection)：表中 TableRow 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="27d8c-133">[TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): A collection of TableRow objects in a Table.</span></span>
 
- <span data-ttu-id="27d8c-134">[TableCell](/javascript/api/onenote/onenote.tablecell)：段落对象的容器。</span><span class="sxs-lookup"><span data-stu-id="27d8c-134">[TableCell](/javascript/api/onenote/onenote.tablecell): A container for Paragraph objects.</span></span>
    - <span data-ttu-id="27d8c-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection)：TableRow 中 TableCell 对象的集合。</span><span class="sxs-lookup"><span data-stu-id="27d8c-135">[TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): A collection of TableCell objects in a TableRow.</span></span>

## <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="27d8c-136">OneNote JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="27d8c-136">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="27d8c-137">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="27d8c-137">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="27d8c-138">Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。</span><span class="sxs-lookup"><span data-stu-id="27d8c-138">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="27d8c-139">有关 OneNote JavaScript API 要求集的详细信息，请参阅 [OneNote JavaScript API 要求集](../requirement-sets/onenote-api-requirement-sets.md)文章。</span><span class="sxs-lookup"><span data-stu-id="27d8c-139">For detailed information about OneNote JavaScript API requirement sets, see the [OneNote JavaScript API requirement sets](../requirement-sets/onenote-api-requirement-sets.md) article.</span></span>

## <a name="onenote-javascript-api-reference"></a><span data-ttu-id="27d8c-140">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="27d8c-140">OneNote JavaScript API reference</span></span>

<span data-ttu-id="27d8c-141">有关 OneNote JavaScript API 的详细信息，请参阅 [OneNote JavaScript API 参考文档](/javascript/api/onenote)。</span><span class="sxs-lookup"><span data-stu-id="27d8c-141">For detailed information about the OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).</span></span>

## <a name="see-also"></a><span data-ttu-id="27d8c-142">另请参阅</span><span class="sxs-lookup"><span data-stu-id="27d8c-142">See also</span></span>

- [<span data-ttu-id="27d8c-143">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="27d8c-143">OneNote JavaScript API programming overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [<span data-ttu-id="27d8c-144">生成第一个 OneNote 外接程序</span><span class="sxs-lookup"><span data-stu-id="27d8c-144">Build your first OneNote add-in</span></span>](../../quickstarts/onenote-quickstart.md)
- [<span data-ttu-id="27d8c-145">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="27d8c-145">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="27d8c-146">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="27d8c-146">Office Add-ins platform overview</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
