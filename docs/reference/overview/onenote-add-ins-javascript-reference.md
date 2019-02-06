---
title: OneNote JavaScript API 概述
description: ''
ms.date: 10/09/2018
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 87bc16f77c14871044fa628f9903ea6ae05f3e0e
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742420"
---
# <a name="onenote-javascript-api-overview"></a>OneNote JavaScript API 概述

适用于：OneNote Online

下面的链接展示了 API 中的高级 OneNote 对象。 每个对象页面链接包含对象可用的属性、事件和方法的描述。 如需了解详细信息，请浏览相应链接。 
    
- [Application](/javascript/api/onenote/onenote.application)：用于访问所有全局可寻址的 OneNote 对象（如活动笔记本和活动分区）的顶级对象。

- [笔记本](/javascript/api/onenote/onenote.notebook)：一个笔记本。笔记本包含分区组合和分区。
    - [NotebookCollection](/javascript/api/onenote/onenote.notebookcollection)：笔记本的集合。

- [SectionGroup](/javascript/api/onenote/onenote.sectiongroup)：一个分区组。分区组包含分区组和分区。
    - [SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection)：分区组的集合。

- [Section](/javascript/api/onenote/onenote.section)：一个分区。分区包含页面。
    - [SectionCollection](/javascript/api/onenote/onenote.sectioncollection)：分区的集合。

- [Page](/javascript/api/onenote/onenote.page)：一个页面。页面包含 PageContent 对象。
    - [PageCollection](/javascript/api/onenote/onenote.pagecollection)：页面的集合。

- [PageContent](/javascript/api/onenote/onenote.pagecontent)：页面上包含内容类型的顶级地区，例如 Outline 或 Image。可在页面上为 PageContent 对象分配一个位置。
    - [PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection)：PageContent 对象的集合，表示页面的内容。

- [Outline](/javascript/api/onenote/onenote.outline)：Paragraph 对象的容器。Outline 是 PageContent 对象的直接子级。

- [Image](/javascript/api/onenote/onenote.image)：Image 对象。Image 可以是 PageContent 对象或 Paragraph 的直接子级。

- [Paragraph](/javascript/api/onenote/onenote.paragraph)：页面上可见内容的容器。Paragraph 是 Outline 的直接子级。
    - [ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection)：Outline 中 Paragraph 对象的集合。

- [RichText](/javascript/api/onenote/onenote.richtext)：RichText 对象。

- [表格](/javascript/api/onenote/onenote.table)：TableRow 对象的容器。

- [TableRow](/javascript/api/onenote/onenote.tablerow)：TableCell 对象的容器。
    - [TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection)：表中 TableRow 对象的集合。
 
- [TableCell](/javascript/api/onenote/onenote.tablecell)：段落对象的容器。
    - [TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection)：TableRow 中 TableCell 对象的集合。

## <a name="onenote-javascript-api-requirement-sets"></a>OneNote JavaScript API 要求集

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。 有关 OneNote JavaScript API 要求集的详细信息，请参阅 [OneNote JavaScript API 要求集](../requirement-sets/onenote-api-requirement-sets.md)文章。

## <a name="onenote-javascript-api-reference"></a>OneNote JavaScript API 参考

有关 OneNote JavaScript API 的详细信息，请参阅 [OneNote JavaScript API 参考文档](/javascript/api/onenote)。

## <a name="see-also"></a>另请参阅

- [OneNote JavaScript API 编程概述](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [生成第一个 OneNote 外接程序](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 加载项平台概述](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
