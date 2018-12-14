---
title: 使用 Office Open XML 创建更优质的 Word 加载项
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: f178a9ee05661e69cc5e08857bbdf8f5081553e0
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27271046"
---
# <a name="create-better-add-ins-for-word-with-office-open-xml"></a><span data-ttu-id="f1c9e-102">使用 Office Open XML 创建更优质的 Word 加载项</span><span class="sxs-lookup"><span data-stu-id="f1c9e-102">Create better add-ins for Word with Office Open XML</span></span>

<span data-ttu-id="f1c9e-103">**提供者：** Stephanie Krieger，Microsoft Corporation | Juan Balmori Labra，Microsoft Corporation</span><span class="sxs-lookup"><span data-stu-id="f1c9e-103">**Provided by:** Stephanie Krieger, Microsoft Corporation | Juan Balmori Labra, Microsoft Corporation</span></span>

<span data-ttu-id="f1c9e-p101">如果您构建在 Word 中运行的 Office 外接程序，则您可能已经了解适用于 Office 的 JavaScript API (Office.js) 提供了多种读取和写入文档内容的格式。这些称为强制类型，包括纯文本、表格、HTML 以及 Office Open XML。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p101">If you're building Office Add-ins to run in Word, you might already know that the JavaScript API for Office (Office.js) offers several formats for reading and writing document content. These are called coercion types, and they include plain text, tables, HTML, and Office Open XML.</span></span>

<span data-ttu-id="f1c9e-p102">因此，当您需要向文档添加多种格式的内容（如图像、格式化表格、图表，甚至仅为格式化文本）时，会进行什么选择？你可以使用 HTML 来插入一些多种格式内容的类型，例如图片。HTML 强制转换可能有一些缺点，例如对内容可用的格式设置和定位选项的限制，具体取决于你的方案。由于 Office Open XML 是用于编写 Word 文档（例如 .docx 和 .dotx）的语言，因此您可以使用用户可以应用的几乎任何类型的格式设置插入用户可以添加到 Word 文档中的几乎任何类型的内容。确定需要完成的 Office Open XML 标记比你想象的容易。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p102">So what are your options when you need to add rich content to a document, such as images, formatted tables, charts, or even just formatted text? You can use HTML for inserting some types of rich content, such as pictures. Depending on your scenario, there can be drawbacks to HTML coercion, such as limitations in the formatting and positioning options available to your content. Because Office Open XML is the language in which Word documents (such as .docx and .dotx) are written, you can insert virtually any type of content that a user can add to a Word document, with virtually any type of formatting the user can apply. Determining the Office Open XML markup you need to get it done is easier than you might think.</span></span>

> [!NOTE]
> <span data-ttu-id="f1c9e-p103">Office Open XML 也是 PowerPoint 和 Excel（以及 Office 2013 及更高版本中的 Visio）文档的技术支持语言。不过，目前只能在 Office Word 加载项中将内容强制转换为 Office Open XML。若要详细了解 Office Open XML（包括完整语言参考文档），请参阅[其他资源](#see-also)。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p103">Office Open XML is also the language behind PowerPoint and Excel (and, as of Office 2013, Visio) documents. However, currently, you can coerce content as Office Open XML only in Office Add-ins created for Word. For more information about Office Open XML, including the complete language reference documentation, see [Additional resources](#see-also).</span></span>

<span data-ttu-id="f1c9e-p104">开始之前，请查看可以使用 Office Open XML 强制转换插入的内容类型。下载代码示例 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML)，其中包含在 Word 中插入以下任何示例所需的 Office Open XML 标记和 Office.js 代码。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p104">To begin, take a look at some of the content types you can insert using Office Open XML coercion. Download the code sample [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), which contains the Office Open XML markup and Office.js code required for inserting any of the following examples into Word.</span></span>

> [!NOTE]
> <span data-ttu-id="f1c9e-116">本文通篇使用的术语**内容类型**和**丰富内容**是指可以插入 Word 文档的丰富内容类型。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-116">Throughout this article, the terms  **content types** and **rich content** refer to the types of rich content you can insert into a Word document.</span></span>


<span data-ttu-id="f1c9e-117">*图 1：应用了直接格式的文本*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-117">*Figure 1. Text with direct formatting*</span></span>


![应用了直接格式的文本。](../images/office15-app-create-wd-app-using-ooxml-fig01.png)

<span data-ttu-id="f1c9e-119">无论用户文档中的现有格式如何，都可以使用直接格式精确指定文本的外观。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-119">You can use direct formatting to specify exactly what the text will look like regardless of existing formatting in the user's document.</span></span>

<span data-ttu-id="f1c9e-120">*图 2：使用样式格式化的文本*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-120">*Figure 2. Text formatted using a style*</span></span>


![使用段落样式格式化的文本。](../images/office15-app-create-wd-app-using-ooxml-fig02.png)

<span data-ttu-id="f1c9e-122">可以使用样式自动协调插入用户文档的文本的外观。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-122">You can use a style to automatically coordinate the look of text you insert with the user's document.</span></span>

<span data-ttu-id="f1c9e-123">*图 3：简单图像*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-123">*Figure 3. A simple image*</span></span>


![徽标图像。](../images/office15-app-create-wd-app-using-ooxml-fig03.png)

<span data-ttu-id="f1c9e-125">可以使用相同的方法，插入 Office 支持的所有格式图像。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-125">You can use the same method for inserting any Office-supported image format.</span></span>

<span data-ttu-id="f1c9e-126">*图 4：使用图片样式和效果格式化的图像*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-126">*Figure 4. An image formatted using picture styles and effects*</span></span>


![Word 2013 中的格式化图像。](../images/office15-app-create-wd-app-using-ooxml-fig04.png)


<span data-ttu-id="f1c9e-128">向图像应用优质格式和效果所需的标记比预期要少。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-128">Adding high quality formatting and effects to your images requires much less markup than you might expect.</span></span>

<span data-ttu-id="f1c9e-129">*图 5：内容控件*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-129">*Figure 5. A content control*</span></span>


![绑定内容控件中的文本。](../images/office15-app-create-wd-app-using-ooxml-fig05.png)

<span data-ttu-id="f1c9e-131">可以结合使用加载项和内容控件，将内容添加到指定（绑定）位置，而不是随意选择的位置。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-131">You can use content controls with your add-in to add content at a specified (bound) location rather than at the selection.</span></span>

<span data-ttu-id="f1c9e-132">*图 6：应用了艺术字格式的文本框*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-132">*Figure 6. A text box with WordArt formatting*</span></span>


![使用艺术字文本效果格式化的文本。](../images/office15-app-create-wd-app-using-ooxml-fig06.png)

<span data-ttu-id="f1c9e-134">文本效果可用于 Word 中的文本框文本（如此处所示），也可用于常规正文文本。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-134">Text effects are available in Word for text inside a text box (as shown here) or for regular body text.</span></span>

<span data-ttu-id="f1c9e-135">*图 7：形状*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-135">*Figure 7. A shape*</span></span>


![Word 2013 中的 Office 2013 绘图形状。](../images/office15-app-create-wd-app-using-ooxml-fig07.png)

<span data-ttu-id="f1c9e-137">可以插入带/不带文本和格式效果的内置或自定义绘图形状。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-137">You can insert built-in or custom drawing shapes, with or without text and formatting effects.</span></span>

<span data-ttu-id="f1c9e-138">*图 8：应用了直接格式的表格*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-138">*Figure 8. A table with direct formatting*</span></span>


![Word 2013 中的格式化表格。](../images/office15-app-create-wd-app-using-ooxml-fig08.png)

<span data-ttu-id="f1c9e-140">可以包括文本格式、边框、阴影、单元格尺寸调整或所需的任何表格格式。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-140">You can include text formatting, borders, shading, cell sizing, or any table formatting you need.</span></span>

<span data-ttu-id="f1c9e-141">*图 9：使用表格样式格式化的表格*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-141">*Figure 9. A table formatted using a table style*</span></span>


![Word 2013 中的格式化表格。](../images/office15-app-create-wd-app-using-ooxml-fig09.png)

<span data-ttu-id="f1c9e-143">可以使用内置或自定义表格样式，就像对文本使用段落样式一样简单。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-143">You can use built-in or custom table styles just as easily as using a paragraph style for text.</span></span>

<span data-ttu-id="f1c9e-144">*图 10：SmartArt 图表*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-144">*Figure 10. A SmartArt diagram*</span></span>


![Word 2013 中的动态 SmartArt 图表。](../images/office15-app-create-wd-app-using-ooxml-fig10.png)

<span data-ttu-id="f1c9e-146">Office 2013 提供了大量 SmartArt 图表布局（可以使用 Office Open XML 创建自己的 SmartArt 图表布局）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-146">Office 2013 offers a wide array of SmartArt diagram layouts (and you can use Office Open XML to create your own).</span></span>

<span data-ttu-id="f1c9e-147">*图 11：图表*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-147">*Figure 11. A chart*</span></span>


![Word 2013 中的图表。](../images/office15-app-create-wd-app-using-ooxml-fig11.png)

<span data-ttu-id="f1c9e-p105">你可以在 Word 文档中插入 Excel 图表作为实时图表，这也意味着你可以在 Word 外接程序中使用这些图表。如上述示例中所示，你可以使用 Office Open XML 强制转换，以插入用户可以插入其自己的文档中的几乎任何类型的内容。获取所需的 Office Open XML 标记有两种简单的方法。将多种格式的内容添加到一个原本空白的 Word 2013 文档中，然后将文件保存为 Word XML 文档格式，或通过 [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) 方法，使用测试外接程序来捕捉标记。两种方法都可以获得几乎相同的结果。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p105">You can insert Excel charts as live charts in Word documents, which also means you can use them in your add-in for Word. As you can see by the preceding examples, you can use Office Open XML coercion to insert essentially any type of content that a user can insert into their own document. There are two simple ways to get theOffice Open XML markup you need. Either add your rich content to an otherwise blank Word 2013 document and then save the file in Word XML Document format or use a test add-in with the [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method to grab the markup. Both approaches provide essentially the same result.</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-p106">Office Open XML 文档实际上是表示文档内容的文件压缩包。以 Word XML 文档格式保存文件可获得整个 Office Open XML 包（合并到一个 XML 文件中），也可以在使用 **getSelectedDataAsync** 检索 Office Open XML 标记时获取相同的包。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p106">An Office Open XML document is actually a compressed package of files that represent the document contents. Saving the file in the Word XML Document format gives you the entireOffice Open XML package flattened into one XML file, which is also what you get when using  **getSelectedDataAsync** to retrieve the Office Open XML markup.</span></span>

<span data-ttu-id="f1c9e-p107">如果从 Word 中将文件保存为 XML 格式，请注意“另存为”对话框的“另存为类型”列表下有两个适用于 .xml 格式文件的选项。请务必选择“**Word XML 文档**”，而非 Word 2003 选项。下载名为 [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) 的代码示例，该示例可以用作检索和测试标记的工具。这就是全部内容吗？并不完全是。是的，对于很多方案而言，你可以使用通过上述任意方法得到的完整的平展 Office Open XML 结果，且其可行。好消息是，你可能无需大部分标记。如果你是首次看到 Office Open XML 标记的众多外接程序开发人员之一，尝试了解为最简单的内容获取的大量标记可能会令人不知所措，但无需如此。在本主题中，我们将使用从 Office 外接程序开发人员社区听到的一些常见方案向你展示用于简化 Office Open XML 以便在外接程序中使用的技术。我们将探讨针对之前所述的部分类型的内容的标记以及最大限度减少 Office Open XML 负载所需的信息。我们还会介绍将多种格式的内容插入文档的活动选择区时所需的代码，以及如何将 Office Open XML 与绑定对象结合使用以在指定位置添加或替换内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p107">If you save the file to an XML format from Word, note that there are two options under the Save as Type list in the Save As dialog box for .xml format files. Be sure to choose  **Word XML Document** and not the Word 2003 option. Download the code sample named [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), which you can use as a tool to retrieve and test your markup. So is that all there is to it? Well, not quite. Yes, for many scenarios, you could use the full, flattened Office Open XML result you see with either of the preceding methods and it would work. The good news is that you probably don't need most of that markup. If you're one of the many add-in developers seeing Office Open XML markup for the first time, trying to make sense of the massive amount of markup you get for the simplest piece of content might seem overwhelming, but it doesn't have to be. In this topic, we'll use some common scenarios we've been hearing from the Office Add-ins developer community to show you techniques for simplifying Office Open XML for use in your add-in. We'll explore the markup for some types of content shown earlier along with the information you need for minimizing the Office Open XML payload. We'll also look at the code you need for inserting rich content into a document at the active selection and how to use Office Open XML with the bindings object to add or replace content at specified locations.</span></span>

## <a name="exploring-the-office-open-xml-document-package"></a><span data-ttu-id="f1c9e-167">探讨 Office Open XML 文档包</span><span class="sxs-lookup"><span data-stu-id="f1c9e-167">Exploring the Office Open XML document package</span></span>


<span data-ttu-id="f1c9e-p108">在使用 [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) 检索选定内容的 Office Open XML 时（或在将文档保存为 Word XML 文档格式时），获取的内容不仅仅是描述选定内容的标记；它是带有您几乎肯定不需要的多个选项和设置的整个文档。事实上，如果对包含任务窗格外接程序的文档使用此方法，则获取的标记甚至包括您的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p108">When you use [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) to retrieve the Office Open XML for a selection of content (or when you save the document in Word XML Document format), what you're getting is not just the markup that describes your selected content; it's an entire document with many options and settings that you almost certainly don't need. In fact, if you use that method from a document that contains a task pane add-in, the markup you get even includes your task pane.</span></span>

<span data-ttu-id="f1c9e-170">即使是简单的 Word 文档包，除了实际内容的部件之外，还包括文档属性、样式、主题（格式设置）、Web 设置、字体等的部件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-170">Even a simple Word document package includes parts for document properties, styles, theme (formatting settings), web settings, fonts, and then some, in addition to parts for the actual content.</span></span>

<span data-ttu-id="f1c9e-p109">例如，假设您只想要插入直接格式的文本段落，如前面图 1 中所示。在使用 **getSelectedDataAsync** 捕捉 Office Open XML 的格式化文本时，可以看到大量标记。这些标记包括表示整个文档的数据包元素，其中包含多个部件（通常称为文档部件，在 Office Open XML 中称为数据包部件），如图 13 中所示。每个部件表示数据包中的一个单独文件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p109">For example, say that you want to insert just a paragraph of text with direct formatting, as shown earlier in Figure 1. When you grab the Office Open XML for the formatted text using  **getSelectedDataAsync**, you see a large amount of markup. That markup includes a package element that represents an entire document, which contains several parts (commonly referred to as document parts or, in the Office Open XML, as package parts), as you see listed in Figure 13. Each part represents a separate file within the package.</span></span>


> [!TIP]
> <span data-ttu-id="f1c9e-p110">可以在记事本等文本编辑器中编辑 Office Open XML 标记。如果在 Visual Studio 2015 中打开它，可以使用“编辑 > 高级 > 格式化文档”\*\*\*\*（Ctrl+K、Ctrl+D）设置包格式，以简化编辑。然后，可以折叠或展开其中的文档部分，如图 12 所示，以便更轻松地查看和编辑 Office Open XML 包内容。每个文档部分都是以 **pkg:part** 标记开头。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p110">You can edit Office Open XML markup in a text editor like Notepad. If you open it in Visual Studio 2015, you can use  **Edit >Advanced > Format Document** (Ctrl+K, Ctrl+D) to format the package for easier editing. Then you can collapse or expand document parts or sections of them, as shown in Figure 12, to more easily review and edit the content of the Office Open XML package. Each document part begins with a **pkg:part** tag.</span></span>


<span data-ttu-id="f1c9e-179">*图 12：折叠和展开包部分以便在 Visual Studio 2015 中更轻松地编辑*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-179">*Figure 12. Collapse and expand package parts for easier editing in Visual Studio 2015*</span></span>

![包部件的 Office Open XML 代码段。](../images/office15-app-create-wd-app-using-ooxml-fig12.png)

<span data-ttu-id="f1c9e-181">*图 13：基本 Word Office Open XML 文档包中的各部分*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-181">*Figure 13. The parts included in a basic Word Office Open XML document package*</span></span>

![包部件的 Office Open XML 代码段。](../images/office15-app-create-wd-app-using-ooxml-fig13.png)

<span data-ttu-id="f1c9e-183">通过所有标记，您会惊奇地发现您真正需要插入格式化文本示例的元素就是 .rels 部件和 document.xml 部件的片段。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-183">With all that markup, you might be surprised to discover that the only elements you actually need to insert the formatted text example are pieces of the .rels part and the document.xml part.</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-p111">包标记上方有两行标记（版本 XML 声明和 Office 程序 ID）的前提是，使用 Office Open XML 强制转换类型，因此无需将它们包括在内。若要将编辑过的标记打开为 Word 文档以进行测试，请保留这两行标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p111">The two lines of markup above the package tag (the XML declarations for version and Office program ID) are assumed when you use the Office Open XML coercion type, so you don't need to include them. Keep them if you want to open your edited markup as a Word document to test it.</span></span>

<span data-ttu-id="f1c9e-p112">本主题开始介绍的多个其他类型的内容也需要其他部件（图 13 中所示之外的部件），我们将在本主题中稍后介绍。同时，您将看到图 13 中所示的任何 Word 文档包标记中的大部分部件，因此此处有一个关于每个部件的作用以及何时需要这些部件的快速摘要。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p112">Several of the other types of content shown at the start of this topic require additional parts as well (beyond those shown in Figure 13), and we'll address those later in this topic. Meanwhile, since you'll see most of the parts shown in Figure 13 in the markup for any Word document package, here's a quick summary of what each of these parts is for and when you need it:</span></span>


- <span data-ttu-id="f1c9e-p113">数据包标记内部的第一个部件是 .rels 文件，它定义数据包顶级各部件之间的关系（通常为文档属性、缩略图(如果有)，以及主文档正文）。标记中始终需要部件中的一些内容，因为您需要将（内容所在的）主文档部件的关系定义为文档包。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p113">Inside the package tag, the first part is the .rels file, which defines relationships between the top-level parts of the package (these are typically the document properties, thumbnail (if any), and main document body). Some of the content in this part is always required in your markup because you need to define the relationship of the main document part (where your content resides) to the document package.</span></span>

- <span data-ttu-id="f1c9e-190">document.xml.rels 部分定义了 document.xml（正文）部分（若有）所需的其他部分的关系。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-190">The document.xml.rels part defines relationships for additional parts required by the document.xml (main body) part, if any.</span></span>


   > [!IMPORTANT]
   > <span data-ttu-id="f1c9e-p114">数据包（如顶级 .rels、document.xml.rels 以及其他可以看到的特定内容类型的数据包）中的 .rels 文件是一个非常重要的工具，您可以将其作为指南，帮助您快速编辑 Office Open XML 数据包。若要了解有关详细信息，请参阅本主题后面的[创建您自己的标记：最佳做法](#creating-your-own-markup-best-practices)。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p114">The .rels files in your package (such as the top-level .rels, document.xml.rels, and others you may see for specific types of content) are an extremely important tool that you can use as a guide for helping you quickly edit down your Office Open XML package. To learn more about how to do this, see [Creating your own markup: best practices](#creating-your-own-markup-best-practices) later in this topic.</span></span>



- <span data-ttu-id="f1c9e-p115">document.xml 部件是文档正文中的内容。您当然需要此部件的元素，因为它正好是内容显示的位置。但您并不需要在此部件中看到的所有内容。我们将在后面详细介绍。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p115">The document.xml part is the content in the main body of the document. You need elements of this part, of course, since that's where your content appears. But, you don't need everything you see in this part. We'll look at that in more detail later.</span></span>

- <span data-ttu-id="f1c9e-p116">在使用 Office Open XML 强制转换将内容插入到文档中时，很多部件会自动被 Set 方法忽略，因此您可能还要删除它们。这些部件包括 theme1.xml 文件（文档的格式主题）、文档属性部件（核心、外接程序和缩略图），以及设置文件（包括设置、webSettings 和 fontTable）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p116">Many parts are automatically ignored by the Set methods when inserting content into a document using Office Open XML coercion, so you might as well remove them. These include the theme1.xml file (the document's formatting theme), the document properties parts (core, add-in, and thumbnail), and setting files (including settings, webSettings, and fontTable).</span></span>

- <span data-ttu-id="f1c9e-p117">在图 1 示例中，直接应用文本格式（即单独应用每个字体和段落格式设置）。但是，如果按前面的图 2 中所示使用样式（例如，如果您想让文本在目标文档中自动呈现“Heading 1”样式），则您可能需要部分 styles.xml 部件及其关系定义。有关详细信息，请参阅主题节“[添加使用其他 Office Open XML 部件的对象](#adding-objects-that-use-additional-office-open-xml-parts)”。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p117">In the Figure 1 example, text formatting is directly applied (that is, each font and paragraph formatting setting applied individually). But, if you use a style (such as if you want your text to automatically take on the formatting of the Heading 1 style in the destination document) as shown earlier in Figure 2, then you would need part of the styles.xml part as well as a relationship definition for it. For more information, see the topic section [Adding objects that use additional Office Open XML parts](#adding-objects-that-use-additional-office-open-xml-parts).</span></span>


## <a name="inserting-document-content-at-the-selection"></a><span data-ttu-id="f1c9e-202">在选定内容插入文档内容</span><span class="sxs-lookup"><span data-stu-id="f1c9e-202">Inserting document content at the selection</span></span>


<span data-ttu-id="f1c9e-203">我们来看看图 1 中所示的格式化文本示例所需的最少的 Office Open XML 标记，以及在文档的活动选定区插入此标记所需的 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-203">Let's take a look at the minimal Office Open XML markup required for the formatted text example shown in Figure 1 and the JavaScript required for inserting it at the active selection in the document.</span></span>


### <a name="simplified-office-open-xml-markup"></a><span data-ttu-id="f1c9e-204">简化的 Office Open XML 标记</span><span class="sxs-lookup"><span data-stu-id="f1c9e-204">Simplified Office Open XML markup</span></span>

<span data-ttu-id="f1c9e-p118">如上文所述，我们已经编辑了此处所示的 Office Open XML 示例，以仅保留需要的文档部件以及每个部件中需要的元素。我们将在本主题的下一节中逐步完成自己编辑标记的步骤（并对此处遗留的片段作进一步解释）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p118">We've edited the Office Open XML example shown here, as described in the preceding section, to leave just required document parts and only required elements within each of those parts. We'll walk through how to edit the markup yourself (and explain a bit more about the pieces that remain here) in the next section of the topic.</span></span>


```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
```


> [!NOTE]
> <span data-ttu-id="f1c9e-p119">如果将此处所示的标记与 version XML 声明标记和 mso-application 一起添加到 XML 文件（如图 13 所示，后两行标记位于文件顶部），可以在 Word 中将它打开为 Word 文档。如果没有添加后两行标记，也仍可以通过依次单击 Word 中的“文件”>“打开”\*\*\*\* 打开它。此时，Word 2013 标题栏上显示“兼容性模式”\*\*\*\*，因为已删除指示 Word 这是 2013 文档的设置。由于要将此标记添加到现有 Word 2013 文档，因此内容完全不会受影响。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p119">If you add the markup shown here to an XML file along with the XML declaration tags for version and mso-application at the top of the file (shown in Figure 13), you can open it in Word as a Word document. Or, without those tags, you can still open it using  **File> Open** in Word. You'll see **Compatibility Mode** on the title bar in Word 2013, because you removed the settings that tell Word this is a 2013 document. Since you're adding this markup to an existing Word 2013 document, that won't affect your content at all.</span></span>


### <a name="javascript-for-using-setselecteddataasync"></a><span data-ttu-id="f1c9e-211">使用 setSelectedDataAsync 所需的 JavaScript</span><span class="sxs-lookup"><span data-stu-id="f1c9e-211">JavaScript for using setSelectedDataAsync</span></span>


<span data-ttu-id="f1c9e-212">将前面的 Office Open XML 保存为解决方案可访问的 XML 文件后，就可以使用以下函数设置使用 Office Open XML 强制转换的文档中的格式化文本内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-212">Once you save the preceding Office Open XML as an XML file that's accessible from your solution, you can use the following function to set the formatted text content in the document using Office Open XML coercion.</span></span> 

<span data-ttu-id="f1c9e-p120">在此函数中，请注意除了最后一行，其他都用于获取已保存的标记，以用于函数末尾的 [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) 方法调用。**setSelectedDataASync** 仅要求您指定要插入的内容以及强制类型。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p120">In this function, notice that all but the last line are used to get your saved markup for use in the [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) method call at the end of the function. **setSelectedDataASync** requires only that you specify the content to be inserted and the coercion type.</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-p121">将 _yourXMLfilename_ 替换为在解决方案中保存的 XML 文件的名称和路径。如果不确定将 XML 文件保存到解决方案中的哪个位置，或不确定如何在代码中进行引用，请参阅 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 代码示例查看相关示例，以及本文展示的有效标记和 JavaScript 示例。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p121">Replace  _yourXMLfilename_ with the name and path of the XML file as you've saved it in your solution. If you're not sure where to include XML files in your solution or how to reference them in your code, see the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample for examples of that and a working example of the markup and JavaScript shown here.</span></span>




```js
function writeContent() {
    var myOOXMLRequest = new XMLHttpRequest();
    var myXML;
    myOOXMLRequest.open('GET', 'yourXMLfilename', false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });
}
```


## <a name="creating-your-own-markup-best-practices"></a><span data-ttu-id="f1c9e-217">创建自己的标记：最佳做法</span><span class="sxs-lookup"><span data-stu-id="f1c9e-217">Creating your own markup: best practices</span></span>


<span data-ttu-id="f1c9e-218">让我们来仔细看看您插入前面的格式化文本示例需要的标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-218">Let's take a closer look at the markup you need to insert the preceding formatted text example.</span></span>

<span data-ttu-id="f1c9e-p122">对于此示例，首先只需从数据包（而不是 .rels 和 document.xml）中删除所有文档部件。然后，我们将编辑两个必需的部件以进一步简化。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p122">For this example, start by simply deleting all document parts from the package other than .rels and document.xml. Then, we'll edit those two required parts to simplify things further.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="f1c9e-p123">请将 .rels 部分用作地图，以快速判断包中内容，并确定可以完全删除的部分（即与内容不相关或内容未引用的任何部分）。请注意，必须在包中定义每个文档部分的关系，这些关系显示在 .rels 文件中。因此，应该能够看到所有关系在 .rels、document.xml.rels 或内容专用 .rels 文件中列出。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p123">Use the .rels parts as a map to quickly gauge what's included in the package and determine what parts you can delete completely (that is, any parts not related to or referenced by your content). Remember that every document part must have a relationship defined in the package and those relationships appear in the .rels files. So you should see all of them listed in either .rels, document.xml.rels, or a content-specific .rels file.</span></span>

<span data-ttu-id="f1c9e-p124">以下标记说明了编辑之前所需的 .rels 部件。我们删除的是外接程序和核心文档属性部件，以及缩略图部件，因此还需要从 .rels 删除这些关系。请注意，这将仅保留 document.xml 的关系（和以下示例中的关系 ID“rID1”）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p124">The following markup shows the required .rels part before editing. Since we're deleting the add-in and core document property parts, and the thumbnail part, we need to delete those relationships from .rels as well. Notice that this will leave only the relationship (with the relationship ID "rID1" in the following example) for document.xml.</span></span>




```XML
<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
  <pkg:xmlData>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.emf"/>
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    </Relationships>
  </pkg:xmlData>
</pkg:part>
```


> [!IMPORTANT]
> <span data-ttu-id="f1c9e-p125">删除从包中完全删除的任何部分的关系（即 **Relationship** 标记）。无论是添加没有定义相应关系的部分，还是删除部分但其关系保留在包中，都会导致错误发生。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p125">Remove the relationships (that is, the **Relationship** tag) for any parts that you completely remove from the package. Including a part without a corresponding relationship, or excluding a part and leaving its relationship in the package, will result in an error.</span></span>

<span data-ttu-id="f1c9e-229">下面的标记展示了 document.xml 部分，其中包括编辑前的示例格式化文本内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-229">The following markup shows the document.xml part, which includes our sample formatted text content before editing.</span></span>

```XML
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document mc:Ignorable="w14 w15 wp14" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <w:body>
          <w:p>
            <w:pPr>
              <w:spacing w:before="360" w:after="0" w:line="480" w:lineRule="auto"/>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:color w:val="70AD47" w:themeColor="accent6"/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="_GoBack"/>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
          <w:p/>
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
</pkg:part>
```

<span data-ttu-id="f1c9e-p126">由于 document.xml 是您放置内容的主要文档部件，我们来快速了解一下此部件的各个方面。（列表后的图 14 提供了可视化参考，说明此处解释的一些相关核心内容和格式标记如何与您在 Word 文档中看到的内容关联。）</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p126">Since document.xml is the primary document part where you place your content, let's take a quick walk through that part. (Figure 14, which follows this list, provides a visual reference to show how some of the core content and formatting tags explained here relate to what you see in a Word document.)</span></span>


- <span data-ttu-id="f1c9e-p127">打开的 **w:document** 标记包括若干个命名空间 (**xmlns**) 列表。其中许多命名空间指的是特定类型的内容，您仅在它们与内容相关时才需要它们。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p127">The opening **w:document** tag includes several namespace ( **xmlns** ) listings. Many of those namespaces refer to specific types of content and you only need them if they're relevant to your content.</span></span>

    <span data-ttu-id="f1c9e-p128">请注意，文档部分的标记前缀引用回命名空间。在本示例中，整个 document.xml 部分的标记中仅使用的前缀为 **w:**，因此，我们需要在 **w:document** 开头标记中保留的唯一命名空间为 **xmlns:w**。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p128">Notice that the prefix for the tags throughout a document part refers back to the namespaces. In this example, the only prefix used in the tags throughout the document.xml part is  **w:**, so the only namespace that we need to leave in the opening **w:document** tag is **xmlns:w**.</span></span>


> [!TIP]
> <span data-ttu-id="f1c9e-p129">若要在 Visual Studio 2015 中编辑标记，请在删除任何部分中的命名空间后，仔细检查相应部分的所有标记。如果删除的是标记的必需命名空间，受影响标记的相关前缀下面会显示红色的弯曲下划线。如果删除 **xmlns:mc** 命名空间，还必须删除命名空间列表前面的 **mc:Ignorable** 属性。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p129">If you're editing your markup in Visual Studio 2015, after you delete namespaces in any part, look through all tags of that part. If you've removed a namespace that's required for your markup, you'll see a red squiggly underline on the relevant prefix for affected tags. If you remove the **xmlns:mc** namespace, you must also remove the **mc:Ignorable** attribute that precedes the namespace listings.</span></span>


- <span data-ttu-id="f1c9e-239">可以在打开的正文标记内看到段落标记 (**w:p**)，其中包含此示例的内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-239">Inside the opening body tag, you see a paragraph tag ( **w:p** ), which includes our sample content for this example.</span></span>

- <span data-ttu-id="f1c9e-p130">**w:pPr** 标记包括直接应用的段落格式的属性，如段落之前或之后的空格、段落对齐方式或缩进。（直接格式指单独应用于内容（而不是作为样式的一部分）的属性。）此标记还包括应用于整个段落的直接字体格式，在嵌套 **w:rPr**（run 属性）标记中，它包括示例中的字体颜色和大小设置。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p130">The **w:pPr** tag includes properties for directly-applied paragraph formatting, such as space before or after the paragraph, paragraph alignment, or indents. (Direct formatting refers to attributes that you apply individually to content rather than as part of a style.) This tag also includes direct font formatting that's applied to the entire paragraph, in a nested **w:rPr** (run properties) tag, which contains the font color and size set in our sample.</span></span>


   > [!NOTE]
   > <span data-ttu-id="f1c9e-p131">可能会注意到，Word Office Open XML 标记中的字号和其他一些格式设置看起来是实际大小的两倍。这是因为段落和行间距以及前面标记所示的一些部分格式属性以缇为单位（磅的二十分之一）。可能还会看到其他多个度量单位，包括用于一些 Office 艺术字 (drawingML) 值的英制单位（914,400 EMU 等于 1 英寸），以及在 drawingML 和 PowerPoint 标记中使用的 100,000 倍实际值，具体视要在 Office Open XML 中使用的内容类型而定。PowerPoint 还将某些值表示为实际值的 100 倍，而 Excel 则通常使用实际值。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p131">You might notice that font sizes and some other formatting settings in Word Office Open XML markup look like they're double the actual size. That's because paragraph and line spacing, as well some section formatting properties shown in the preceding markup, are specified in twips (one-twentieth of a point). Depending on the types of content you work with in Office Open XML, you may see several additional units of measure, including English Metric Units (914,400 EMUs to an inch), which are used for some Office Art (drawingML) values and 100,000 times actual value, which is used in both drawingML and PowerPoint markup. PowerPoint also expresses some values as 100 times actual and Excel commonly uses actual values.</span></span>


- <span data-ttu-id="f1c9e-p132">段落中任何具有相似属性的内容都包括在运行 (**w:r**) 中，如示例文本中的情况。每次格式或内容类型发生更改时，就开始新的运行。（也就是说，如果示例文本中只有一个字是粗体，将会分离到自己的运行中。）本示例中的内容仅包括这一个文本运行。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p132">Within a paragraph, any content with like properties is included in a run ( **w:r** ), such as is the case with the sample text. Each time there's a change in formatting or content type, a new run starts. (That is, if just one word in the sample text was bold, it would be separated into its own run.) In this example, the content includes just the one text run.</span></span>

    <span data-ttu-id="f1c9e-249">请注意，由于本示例中包括的格式是字体格式（即可以应用于一个字符的格式），它还在单独的 run 属性中显示。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-249">Notice that, because the formatting included in this sample is font formatting (that is, formatting that can be applied to as little as one character), it also appears in the properties for the individual run.</span></span>

- <span data-ttu-id="f1c9e-p133">还要注意，隐藏的“_GoBack”书签（**w:bookmarkStart** 和 **w:bookmarkEnd**）的标记默认显示在 Word 2013 文档中。始终可以从标记中删除 GoBack 书签的起始标记和结束标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p133">Also notice the tags for the hidden "_GoBack" bookmark (**w:bookmarkStart** and **w:bookmarkEnd** ), which appear in Word 2013 documents by default. You can always delete the start and end tags for the GoBack bookmark from your markup.</span></span>

- <span data-ttu-id="f1c9e-p134">文档正文的最后一部分是 **w:sectPr** 标记或分节属性。此标记包括边距和页面方向等设置。您使用 **setSelectedDataAsync** 插入的内容将默认呈现在目标文档中的活动部分属性上。因此，除非您的内容包括分节符（能看到多个 **w:sectPr** 标记），否则无法删除此标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p134">The last piece of the document body is the **w:sectPr** tag, or section properties. This tag includes settings such as margins and page orientation. The content you insert using **setSelectedDataAsync** will take on the active section properties in the destination document by default. So, unless your content includes a section break (in which case you'll see more than one **w:sectPr** tag), you can delete this tag.</span></span>


<span data-ttu-id="f1c9e-256">*图 14：document.xml 中的常见标记与 Word 文档内容和布局的对应关系*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-256">*Figure 14. How common tags in document.xml relate to the content and layout of a Word document*</span></span>

![Word 文档中的 Office Open XML 元素。](../images/office15-app-create-wd-app-using-ooxml-fig14.png)

> [!TIP]
> <span data-ttu-id="f1c9e-p135">在创建的标记中，可能还会看到多个标记中有另一个属性，其中包含字符 **w:rsid**（本主题使用的示例中没有此属性）。这些是修订标识符，用于 Word 中的“合并文档”功能，且默认处于启用状态。使用加载项插入标记时，无需使用这些标识符，可以禁用它们，从而简化标记。既能轻松删除现有 RSID 标记，也能禁用此功能（如以下过程所述），这样就不会向新内容的标记添加这些标识符了。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p135">In markup you create, you might see another attribute in several tags that includes the characters **w:rsid**, which you don't see in the examples used in this topic. These are revision identifiers. They're used in Word for the Combine Documents feature and they're on by default. You'll never need them in markup you're inserting with your add-in and turning them off makes for much cleaner markup. You can easily remove existing RSID tags or disable the feature (as described in the following procedure) so that they're not added to your markup for new content.</span></span>

<span data-ttu-id="f1c9e-263">请注意，如果您在 Word 中使用“共同创作”功能（例如与他人同时编辑文档的功能），应在为外接程序完成生成标记后，再次启用此功能。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-263">Be aware that if you use the co-authoring capabilities in Word (such as the ability to simultaneously edit documents with others), you should enable the feature again when finished generating the markup for your add-in.</span></span>

<span data-ttu-id="f1c9e-264">要在 Word 中关闭你创建的文档的 RSID 属性，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-264">To turn off RSID attributes in Word for documents you create going forward, do the following:</span></span> 

1. <span data-ttu-id="f1c9e-265">在 Word 2013 中，选择“**文件**”，然后选择“**选项**”。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-265">In Word 2013, choose **File** and then choose **Options**.</span></span>
2. <span data-ttu-id="f1c9e-266">在“Word 选项”对话框中，选择“**信任中心**”，然后选择“**信任中心设置**”。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-266">In the Word Options dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>
3. <span data-ttu-id="f1c9e-267">在“信任中心”对话框中，选择“**隐私选项**”，然后禁用“**存储随机数以提高组合精确性**”设置。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-267">In the Trust Center dialog box, choose **Privacy Options** and then disable the setting **Store Random Number to Improve Combine Accuracy**.</span></span>

<span data-ttu-id="f1c9e-268">若要从现有文档中删除 RSID 标记，请尝试在 Office Open XML 中打开的文档中使用以下快捷方式：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-268">To remove RSID tags from an existing document, try the following shortcut with the document open in Office Open XML:</span></span>


1. <span data-ttu-id="f1c9e-269">在文档正文中的插入点处按 **Ctrl+Home** 转到文档顶端。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-269">With your insertion point in the main body of the document, press **Ctrl+Home** to go to the top of the document.</span></span>
2. <span data-ttu-id="f1c9e-p136">在键盘上依次按“**空格**”、“**Delete**”、“**空格**”。然后保存文档。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p136">On the keyboard, press **Spacebar**, **Delete**, **Spacebar**. Then, save the document.</span></span>

<span data-ttu-id="f1c9e-272">从此数据包中删除了大部分标记后，只剩下需要为示例插入的最少标记，如上一节中所述。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-272">After removing the majority of the markup from this package, we're left with the minimal markup that needs to be inserted for the sample, as shown in the preceding section.</span></span>


## <a name="using-the-same-office-open-xml-structure-for-different-content-types"></a><span data-ttu-id="f1c9e-273">针对不同内容类型使用相同的 Office Open XML 结构</span><span class="sxs-lookup"><span data-stu-id="f1c9e-273">Using the same Office Open XML structure for different content types</span></span>


<span data-ttu-id="f1c9e-p137">几种类型的多种格式的内容仅需要前面示例中显示的 .rels 和 document.xml 组件，包括内容控件、Office 绘图形状、文本框及表格（除非将样式应用于表格）。事实上，您可以重用已编辑过的相同数据包部件，并仅为内容标记置换出 document.xml 中的 **body** 内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p137">Several types of rich content require only the .rels and document.xml components shown in the preceding example, including content controls, Office drawing shapes and text boxes, and tables (unless a style is applied to the table). In fact, you can reuse the same edited package parts and swap out just the **body** content in document.xml for the markup of your content.</span></span>

<span data-ttu-id="f1c9e-276">若要查看前面图 5 到图 8 中每个内容类型示例的 Office Open XML 标记，可以浏览“概述”部分中引用的 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 代码示例。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-276">To check out the Office Open XML markup for the examples of each of these content types shown earlier in Figures 5 through 8, explore the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample referenced in the overview section.</span></span>

<span data-ttu-id="f1c9e-277">在继续本主题内容之前，我们来看看几个内容类型要注意的差别，以及如何置换出所需的片段。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-277">Before we move on, let's take a look at differences to note for a couple of these content types and how to swap out the pieces you need.</span></span>


### <a name="understanding-drawingml-markup-office-graphics-in-word-what-are-fallbacks"></a><span data-ttu-id="f1c9e-278">了解 Word 中的 drawingML 标记（Office 图形）：什么是回退？</span><span class="sxs-lookup"><span data-stu-id="f1c9e-278">Understanding drawingML markup (Office graphics) in Word: What are fallbacks?</span></span>

<span data-ttu-id="f1c9e-p138">如果形状或文本框的标记看起来要比预期的复杂得多，这是有原因的。我们看到在 Office 2007 版本中引入了 Office Open XML 格式，以及 PowerPoint 和 Excel 完全采用的新 Office 图形引擎。在 2007 版本中，Word 仅并入部分图形引擎，即采用更新的 Excel 图表引擎、SmartArt 图形，以及高级图片工具。对于形状和文本框，Word 2007 继续使用旧的绘图对象 (VML)。Word 在 2010 版本中使用图形引擎执行了其他步骤，以合并更新的图形和绘图工具。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p138">If the markup for your shape or text box looks far more complex than you would expect, there is a reason for it. With the release of Office 2007, we saw the introduction of the Office Open XML Formats as well as the introduction of a new Office graphics engine that PowerPoint and Excel fully adopted. In the 2007 release, Word only incorporated part of that graphics engine, adopting the updated Excel charting engine, SmartArt graphics, and advanced picture tools. For shapes and text boxes, Word 2007 continued to use legacy drawing objects (VML). It was in the 2010 release that Word took the additional steps with the graphics engine to incorporate updated shapes and drawing tools.</span></span>

<span data-ttu-id="f1c9e-284">因此，为了在 Word 2007 中打开 Office Open XML 格式 Word 文档时支持形状和文本框，形状（包括文本框）需要回退 VML 标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-284">So, to support shapes and text boxes in Office Open XML Format Word documents when opened in Word 2007, shapes (including text boxes) require fallback VML markup.</span></span>

<span data-ttu-id="f1c9e-p139">通常情况下，对于 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 代码示例中包括的形状和文本框示例，可以删除回退标记。保存文档后，Word 2013 会自动将缺失的回退标记添加到形状中。但是，如果您更想保留回退标记以确保支持所有用户方案，也不会带来危害。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p139">Typically, as you see for the shape and text box examples included in the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample, the fallback markup can be removed. Word 2013 automatically adds missing fallback markup to shapes when a document is saved. However, if you prefer to keep the fallback markup to ensure that you're supporting all user scenarios, there's no harm in retaining it.</span></span>

<span data-ttu-id="f1c9e-p140">如果内容包括分组绘图对象，您将看到其他（以及明显重复的）标记，但这是必须保留的。当组中包含对象时，绘图形状的标记部分会被复制。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p140">If you have grouped drawing objects included in your content, you'll see additional (and apparently repetitive) markup, but this must be retained. Portions of the markup for drawing shapes are duplicated when the object is included in a group.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="f1c9e-p141">若要使用文本框和绘图形状，请务必先仔细检查命名空间，再将它们从 document.xml 中删除。（或者，若要通过另一个对象类型重用标记，请务必添加回之前可能从 document.xml 中删除的任何必需命名空间。）document.xml 中默认包含的命名空间的重要组成部分旨在满足绘图对象要求。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p141">When working with text boxes and drawing shapes, be sure to check namespaces carefully before removing them from document.xml. (Or, if you're reusing markup from another object type, be sure to add back any required namespaces you might have previously removed from document.xml.) A substantial portion of the namespaces included by default in document.xml are there for drawing object requirements.</span></span>


#### <a name="about-graphic-positioning"></a><span data-ttu-id="f1c9e-292">关于图形位置</span><span class="sxs-lookup"><span data-stu-id="f1c9e-292">About graphic positioning</span></span>

<span data-ttu-id="f1c9e-p142">在代码示例 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 和 [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) 中，使用不同类型的文字环绕和位置设置来设置文本框和形状。（还要注意这些代码示例中的图像示例都根据文本格式进行设置，将图形对象置于文本基线上。）</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p142">In the code samples [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) and [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML), the text box and shape are setup using different types of text wrapping and positioning settings. (Also be aware that the image examples in those code samples are setup using in line with text formatting, which positions a graphic object on the text baseline.)</span></span>

<span data-ttu-id="f1c9e-p143">这些代码示例中的形状的位置相对于页面右边距和下边距进行调整。相对位置可让您更容易协调用户的未知文档设置，因为它将调整用户的边距，并降低由于纸张大小、方向或边距设置而带来的外观突兀的风险。若要在插入图形对象时保留相对位置设置，必须保留存储位置（在 Word 中称为“定位标记”）的段落标记 (w:p)。如果将内容插入现有段落标记，而不是包含自己的标记，您可能可以保留相同的初始可视状态，但很多使位置自动调整用户布局的相对引用类型可能会丢失。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p143">The shape in those code samples is positioned relative to the right and bottom page margins. Relative positioning lets you more easily coordinate with a user's unknown document setup because it will adjust to the user's margins and run less risk of looking awkward because of paper size, orientation, or margin settings. To retain relative positioning settings when you insert a graphic object, you must retain the paragraph mark (w:p) in which the positioning (known in Word as an anchor) is stored. If you insert the content into an existing paragraph mark rather than including your own, you may be able to retain the same initial visual, but many types of relative references that enable the positioning to automatically adjust to the user's layout may be lost.</span></span>


### <a name="working-with-content-controls"></a><span data-ttu-id="f1c9e-299">使用内容控件</span><span class="sxs-lookup"><span data-stu-id="f1c9e-299">Working with content controls</span></span>

<span data-ttu-id="f1c9e-300">内容控件是 Word 2013 中的重要功能，此功能可以通过多种方式大大增强 Word 外接程序的功能，包括使您可以在文档中的指定位置（而不仅仅是选定内容处）插入内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-300">Content controls are an important feature in Word 2013 that can greatly enhance the power of your add-in for Word in multiple ways, including giving you the ability to insert content at designated places in the document rather than only at the selection.</span></span>

<span data-ttu-id="f1c9e-301">在 Word 中，内容控件位于功能区的“开发人员”选项卡上，如图 15 所示。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-301">In Word, find content controls on the Developer tab of the ribbon, as shown here in Figure 15.</span></span>


<span data-ttu-id="f1c9e-302">*图 15：Word 中“开发人员”选项卡上的控件组*</span><span class="sxs-lookup"><span data-stu-id="f1c9e-302">*Figure 15. The Controls group on the Developer tab in Word*</span></span>

![Word 2013 功能区上的内容控件组。](../images/office15-app-create-wd-app-using-ooxml-fig15.png)

<span data-ttu-id="f1c9e-304">Word 中的内容控件类型包括格式文本、纯文本、图片、构建基块库、复选框、下拉列表、组合框、日期选取器，以及重复节。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-304">Types of content controls in Word include rich text, plain text, picture, building block gallery, check box, dropdown list, combo box, date picker, and repeating section.</span></span>



- <span data-ttu-id="f1c9e-305">使用图 15 中所示的“**属性**”命令编辑控件标题，并设置首选项（例如隐藏控件容器）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-305">Use the  **Properties** command, shown in Figure 15, to edit the title of the control and to set preferences such as hiding the control container.</span></span>

- <span data-ttu-id="f1c9e-306">启用“**设计模式**”以编辑控件中的占位符内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-306">Enable  **Design Mode** to edit placeholder content in the control.</span></span>

<span data-ttu-id="f1c9e-p144">如果外接程序使用 Word 模板，则可以在该模板中包含控件，以增强内容行为。你还可以使用 Word 文档中的 XML 数据绑定将内容控件绑定到数据（如文档属性），以轻松完成表单或类似任务。（从“**文档部件**”下的“**插入**”选项卡上可查找已绑定到 Word 中的内置文档属性的控件。）</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p144">If your add-in works with a Word template, you can include controls in that template to enhance the behavior of the content. You can also use XML data binding in a Word document to bind content controls to data, such as document properties, for easy form completion or similar tasks. (Find controls that are already bound to built-in document properties in Word on the  **Insert** tab, under **Quick Parts**.)</span></span>

<span data-ttu-id="f1c9e-p145">您在通过外接程序使用内容控件时，还可以使用不同类型的绑定大幅扩展外接程序可以进行操作的选项。可以从外接程序中绑定内容控件，然后将内容写入到绑定（而不是活动的选定内容）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p145">When you use content controls with your add-in, you can also greatly expand the options for what your add-in can do using a different type of binding. You can bind to a content control from within the add-in and then write content to the binding rather than to the active selection.</span></span>



> [!NOTE]
> <span data-ttu-id="f1c9e-p146">请勿将 Word 中的 XML 数据绑定与通过加载项绑定到控件的功能混淆。它们是完全独立的两种功能。不过，可以将命名内容控件添加到通过加载项使用 OOXML 强制转换插入的内容中，再使用加载项中的代码绑定到这些控件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p146">Don't confuse XML data binding in Word with the ability to bind to a control via your add-in. These are completely separate features. However, you can include named content controls in the content you insert via your add-in using OOXML coercion and then use code in the add-in to bind to those controls.</span></span>

<span data-ttu-id="f1c9e-p147">还要注意 XML 数据绑定和 Office.js 都可以与您应用程序中的自定义 XML 部件交互，因此可以集成这些强大的工具。若要了解有关如何使用 Office JavaScript API 中的自定义 XML 部件的信息，请参阅本主题的[其他资源](#see-also)一节。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p147">Also be aware that both XML data binding and Office.js can interact with custom XML parts in your app, so it is possible to integrate these powerful tools. To learn about working with custom XML parts in the Office JavaScript API, see the [Additional resources](#see-also) section of this topic.</span></span>

<span data-ttu-id="f1c9e-p148">本主题的下一节介绍如何在 Word 外接程序中使用绑定。首先，我们来看看插入可以使用外接程序绑定到的格式文本内容控件所需的 Office Open XML 示例。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p148">Working with bindings in your Word add-in is covered in the next section of the topic. First, let's take a look at an example of the Office Open XML required for inserting a rich text content control that you can bind to using your add-in.</span></span>



> [!IMPORTANT]
> <span data-ttu-id="f1c9e-319">RTF 格式文本控件是可用于在加载项中绑定到内容控件的唯一内容控件类型。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-319">Rich text controls are the only type of content control you can use to bind to a content control from within your add-in.</span></span>




```XML
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" >
        <w:body>
          <w:p/>
          <w:sdt>
              <w:sdtPr>
                <w:alias w:val="MyContentControlTitle"/>
                <w:id w:val="1382295294"/>
                <w15:appearance w15:val="hidden"/>
                <w:showingPlcHdr/>
              </w:sdtPr>
              <w:sdtContent>
                <w:p>
                  <w:r>
                  <w:t>[This text is inside a content control that has its container hidden. You can bind to a content control to add or interact with content at a specified location in the document.]</w:t>
                </w:r>
                </w:p>
              </w:sdtContent>
            </w:sdt>
          </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
 </pkg:package>
```

<span data-ttu-id="f1c9e-320">如前所述，内容控件（如格式化文本）不需要其他文档部件，因此此处仅包含 .rels 和 document.xml 部件的编辑后版本。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-320">As already mentioned, content controls, like formatted text, don't require additional document parts, so only edited versions of the .rels and document.xml parts are included here.</span></span>

<span data-ttu-id="f1c9e-p149">您在 document.xml 正文中看到的 **w:sdt** 标记表示内容控件。如果生成了内容控件的 Office Open XML 标记，则会看到此示例中已删除了多个属性，包括标记和文档部件属性。仅保留了基本的（及几个最佳做法）元素，如下所示：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p149">The **w:sdt** tag that you see within the document.xml body represents the content control. If you generate the Office Open XML markup for a content control, you'll see that several attributes have been removed from this example, including the tag and document part properties. Only essential (and a couple of best practice) elements have been retained, including the following:</span></span>



- <span data-ttu-id="f1c9e-p150">**alias** 是 Word 中“内容控件属性”对话框中的标题属性。如果您计划从外接程序中绑定到控件，则需要此属性（代表项目的名称）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p150">The  **alias** is the title property from the Content Control Properties dialog box in Word. This is a required property (representing the name of the item) if you plan to bind to the control from within your add-in.</span></span>

- <span data-ttu-id="f1c9e-p151">唯一的 **id** 是必需的属性。如果从外接程序中绑定到控件，则 ID 为绑定在文档中用于标识适用的命名内容控件的属性。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p151">The unique **id** is a required property. If you bind to the control from within your add-in, the ID is the property the binding uses in the document to identify the applicable named content control.</span></span>

- <span data-ttu-id="f1c9e-p152">**appearance** 属性用于隐藏控件容器，使外观更简洁。这是 Word 2013 中的一个新功能，通过使用 w15 命名空间可以看到。由于使用了此属性，w15 命名空间会保留在 document.xml 部件的开头。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p152">The  **appearance** attribute is used to hide the control container, for a cleaner look. This is a new feature in Word 2013, as you see by the use of the w15 namespace. Because this property is used, the w15 namespace is retained at the start of the document.xml part.</span></span>

- <span data-ttu-id="f1c9e-p153">**showingPlcHdr** 属性是一个可选的设置，将您包含在控件（此示例中的文本）内的默认内容设置为占位符内容。因此，如果用户在控制区域单击或点按，则选中整个内容，而不是对用户可以更改的可编辑内容进行操作。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p153">The  **showingPlcHdr** attribute is an optional setting that sets the default content you include inside the control (text in this example) as placeholder content. So, if the user clicks or taps in the control area, the entire content is selected rather than behaving like editable content in which the user can make changes.</span></span>

- <span data-ttu-id="f1c9e-p154">尽管 **sdt** 标记前面的空段落标记 (**w:p/**) 不是添加内容控件所必需的（并且将在 Word 文档中的控件上方添加垂直间距），它仍确保控件位于其段落中。它的重要性取决于将在控件中添加的内容的类型和格式。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p154">Although the empty paragraph mark ( **w:p/** ) that precedes the **sdt** tag is not required for adding a content control (and will add vertical space above the control in the Word document), it ensures that the control is placed in its own paragraph. This may be important, depending upon the type and formatting of content that will be added in the control.</span></span>

- <span data-ttu-id="f1c9e-335">如果您想要绑定控件，则控件的默认内容（位于 **sdtContent** 标记中）必须至少包括一个完整的段落（如此示例中所示），以使绑定接受多段落多种格式的内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-335">If you intend to bind to the control, the default content for the control (what's inside the **sdtContent** tag) must include at least one complete paragraph (as in this example), in order for your binding to accept multi-paragraph rich content.</span></span>



> [!NOTE]
> <span data-ttu-id="f1c9e-p155">从此示例 **w:sdt** 标记中删除的文档部分属性可能显示在内容控件中，以在可以存储占位符内容信息的包中引用单独部分（各部分位于 Office Open XML 包的词汇表目录下）。尽管文档部分是用于 Office Open XML 包中 XML 部分（即文件）的术语，sdt 属性中使用的术语“文档部分”是指 Word 中的相同术语，用于描述一些内容类型，包括构建基块和文档属性快速部分（例如，内置 XML 数据绑定控件）。如果在 Office Open XML 包中的词汇表目录下看到部分，可能需要在插入的内容包含这些功能时保留这些部分。对于要在加载项中绑定到的典型内容控件，它们不是必需的。只需注意，如果确实从包中删除词汇表部分，还必须从 w:sdt 标记中删除文档部分属性。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p155">The document part attribute that was removed from this sample **w:sdt** tag may appear in a content control to reference a separate part in the package where placeholder content information can be stored (parts located in a glossary directory in the Office Open XML package). Although document part is the term used for XML parts (that is, files) within an Office Open XML package, the term document parts as used in the sdt property refers to the same term in Word that is used to describe some content types including building blocks and document property quick parts (for example, built-in XML data-bound controls). If you see parts under a glossary directory in your Office Open XML package, you may need to retain them if the content you're inserting includes these features. For a typical content control that you intend to use to bind to from your add-in, they're not required. Just remember that, if you do delete the glossary parts from the package, you must also remove the document part attribute from the w:sdt tag.</span></span>

<span data-ttu-id="f1c9e-341">下一部分将介绍如何在 Word 加载项中创建和使用绑定。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-341">The next section will discuss how to create and use bindings in your Word add-in.</span></span>


## <a name="inserting-content-at-a-designated-location"></a><span data-ttu-id="f1c9e-342">在指定位置插入内容</span><span class="sxs-lookup"><span data-stu-id="f1c9e-342">Inserting content at a designated location</span></span>


<span data-ttu-id="f1c9e-p156">我们已讨论如何在 Word 文档中的活动选定内容处插入内容。如果绑定到文档中的命名内容控件，则可以插入任何同一种内容类型到此控件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p156">We've already looked at how to insert content at the active selection in a Word document. If you bind to a named content control that's in the document, you can insert any of the same content types into that control.</span></span> 

<span data-ttu-id="f1c9e-345">您何时想要使用此方法？</span><span class="sxs-lookup"><span data-stu-id="f1c9e-345">So when might you want to use this approach?</span></span>


- <span data-ttu-id="f1c9e-346">您何时需要在模板中的指定位置添加或替换内容（如从数据库填充文档各个部分）</span><span class="sxs-lookup"><span data-stu-id="f1c9e-346">When you need to add or replace content at specified locations in a template, such as to populate portions of the document from a database</span></span>

- <span data-ttu-id="f1c9e-347">您何时想要替换正插入到活动选定内容处的内容（如为用户提供设计元素选项）的选项</span><span class="sxs-lookup"><span data-stu-id="f1c9e-347">When you want the option to replace content that you're inserting at the active selection, such as to provide design element options to the user</span></span>

- <span data-ttu-id="f1c9e-348">您何时想让用户在文档中添加数据，以使您可以访问并与外接程序一起使用（如根据用户在文档中添加的信息在任务窗格中填充字段）</span><span class="sxs-lookup"><span data-stu-id="f1c9e-348">When you want the user to add data in the document that you can access for use with your add-in, such as to populate fields in the task pane based upon information the user adds in the document</span></span>

<span data-ttu-id="f1c9e-349">下载代码示例 [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings)，此代码示例提供了如何插入并绑定到内容控件及如何填充绑定的可用示例。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-349">Download the code sample [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), which provides a working example of how to insert and bind to a content control, and how to populate the binding.</span></span>


### <a name="add-and-bind-to-a-named-content-control"></a><span data-ttu-id="f1c9e-350">添加并绑定到命名内容控件</span><span class="sxs-lookup"><span data-stu-id="f1c9e-350">Add and bind to a named content control</span></span>


<span data-ttu-id="f1c9e-351">在检查后面的 JavaScript 时，请考虑以下要求：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-351">As you examine the JavaScript that follows, consider these requirements:</span></span>


- <span data-ttu-id="f1c9e-352">如前所述，必须使用富文本控件，以从 Word 外接程序绑定到控件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-352">As previously mentioned, you must use a rich text content control in order to bind to the control from your Word add-in.</span></span>

- <span data-ttu-id="f1c9e-p157">内容控件必须具有名称（这是“内容控件属性”对话框中的“**标题**”字段，对应 Office Open XML 标记中的 **Alias** 标记）。这是代码标识绑定放置位置的方式。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p157">The content control must have a name (this is the  **Title** field in the Content Control Properties dialog box, which corresponds to the **Alias** tag in the Office Open XML markup). This is how the code identifies where to place the binding.</span></span>

- <span data-ttu-id="f1c9e-p158">可以具有多个命名空间，并按需要绑定它们。使用唯一的内容控件名称、唯一的内容控件 ID，以及唯一的绑定 ID。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p158">You can have several named controls and bind to them as needed. Use a unique content control name, unique content control ID, and a unique binding ID.</span></span>


```js
function addAndBindControl() {
    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' }, function (result) {
        if (result.status == "failed") {
            if (result.error.message == "The named item does not exist.")
                var myOOXMLRequest = new XMLHttpRequest();
                var myXML;
                myOOXMLRequest.open('GET', '../../Snippets_BindAndPopulate/ContentControl.xml', false);
                myOOXMLRequest.send();
                if (myOOXMLRequest.status === 200) {
                    myXML = myOOXMLRequest.responseText;
                }
                Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' }, function (result) {
                    Office.context.document.bindings.addFromNamedItemAsync("MyContentControlTitle", "text", { id: 'myBinding' });
                });
        }
    });
}
```

<span data-ttu-id="f1c9e-357">此处所示的代码执行以下步骤：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-357">The code shown here takes the following steps:</span></span>


- <span data-ttu-id="f1c9e-358">尝试使用 [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-) 绑定到命名内容控件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-358">Attempts to bind to the named content control, using [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-).</span></span>

  <span data-ttu-id="f1c9e-p159">如果你的外接程序有可能出现这样一种情况，在执行代码时，文档中已存在命名控件，那么请先执行此步骤。例如，如果外接程序已插入并使用已设计为与该外接程序一起使用的模板进行保存，其中事先放置了该控件，那么你需要执行此操作。如果你需要绑定到该外接程序之前放置的控件，那么你也需要执行此操作。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p159">Take this step first if there is a possible scenario for your add-in where the named control could already exist in the document when the code executes. For example, you'll want to do this if the add-in was inserted into and saved with a template that's been designed to work with the add-in, where the control was placed in advance. You also need to do this if you need to bind to a control that was placed earlier by the add-in.</span></span>

- <span data-ttu-id="f1c9e-p160">对 **addFromNamedItemAsync** 方法首次调用的回退会检查结果状态，以查看绑定是否由于文档中没有命名项目（也就是本示例中名为 MyContentControlTitle 的内容控件）而失败。如果是这样，代码会（使用 **setSelectedDataAsync**）在活动选定内容处添加控件，然后绑定它。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p160">The callback in the first call to the  **addFromNamedItemAsync** method checks the status of the result to see if the binding failed because the named item doesn't exist in the document (that is, the content control named MyContentControlTitle in this example). If so, the code adds the control at the active selection point (using **setSelectedDataAsync** ) and then binds to it.</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-p161">如前所述，以及如前面的代码中所示，内容控件的名称用于确定创建绑定的位置。但是，在 Office Open XML 标记中，代码使用内容控件的名称和 ID 属性添加绑定到文档。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p161">As mentioned earlier and shown in the preceding code, the name of the content control is used to determine where to create the binding. However, in the Office Open XML markup, the code adds the binding to the document using both the name and the ID attribute of the content control.</span></span>

<span data-ttu-id="f1c9e-p162">代码执行之后，如果检查外接程序在其中创建了绑定的文档标记，则会看到每个绑定有两个部件。在（document.xml 中）添加了绑定的内容控件的标记中，会看到 **w15:webExtensionLinked/** 属性。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p162">After code execution, if you examine the markup of the document in which your add-in created bindings, you'll see two parts to each binding. In the markup for the content control where a binding was added (in document.xml), you'll see the attribute  **w15:webExtensionLinked/**.</span></span>

<span data-ttu-id="f1c9e-p163">在名为 webExtensions1.xml 的文档部件中，你将看到已创建的绑定列表。每个绑定都使用绑定 ID 和适用控件的 ID 属性进行标识，如下所示，**appref** 属性为内容控件 ID：\*\* **we:binding id="myBinding" type="text" appref="1382295294"/**。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p163">In the document part named webExtensions1.xml, you'll see a list of the bindings you've created. Each is identified using the binding ID and the ID attribute of the applicable control, such as the following, where the **appref** attribute is the content control ID: \*\* **we:binding id="myBinding" type="text" appref="1382295294"/**.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="f1c9e-p164">必须在要对绑定执行操作时添加绑定。请勿在 Office Open XML 中通过添加绑定标记来插入内容控件，因为插入此标记的过程会删除绑定。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p164">You must add the binding at the time you intend to act upon it. Don't include the markup for the binding in the Office Open XML for inserting the content control because the process of inserting that markup will strip the binding.</span></span>


### <a name="populate-a-binding"></a><span data-ttu-id="f1c9e-372">填充绑定</span><span class="sxs-lookup"><span data-stu-id="f1c9e-372">Populate a binding</span></span>


<span data-ttu-id="f1c9e-373">写入内容到绑定的代码与写入内容到选定内容的代码类似</span><span class="sxs-lookup"><span data-stu-id="f1c9e-373">The code for writing content to a binding is similar to that for writing content to a selection.</span></span>


```js
function populateBinding(filename) {
  var myOOXMLRequest = new XMLHttpRequest();
  var myXML;
  myOOXMLRequest.open('GET', filename, false);
  myOOXMLRequest.send();
  if (myOOXMLRequest.status === 200) {
      myXML = myOOXMLRequest.responseText;
  }
  Office.select("bindings#myBinding").setDataAsync(myXML, { coercionType: 'ooxml' });
}
```

<span data-ttu-id="f1c9e-p165">与 **setSelectedDataAsync** 一样，您可以指定要插入的内容和强制转换类型。写入到绑定的其他唯一要求是通过 ID 标识绑定。请注意此代码 (bindings#myBinding) 中使用的绑定 ID 如何与之前函数创建绑定时建立的绑定 ID (myBinding) 相对应。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p165">As with  **setSelectedDataAsync**, you specify the content to be inserted and the coercion type. The only additional requirement for writing to a binding is to identify the binding by ID. Notice how the binding ID used in this code (bindings#myBinding) corresponds to the binding ID established (myBinding) when the binding was created in the previous function.</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-p166">无论是初始填充绑定，还是替换绑定内容，只需运行上述代码即可。如果在绑定位置插入新的内容片断，相应绑定中的现有内容会自动被替换掉。有关示例，请查看前面引用的代码示例 [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings)，其中提供了两个独立内容示例，可以交替使用它们来填充同一个绑定。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p166">The preceding code is all you need whether you are initially populating or replacing the content in a binding. When you insert a new piece of content at a bound location, the existing content in that binding is automatically replaced. Check out an example of this in the previously-referenced code sample [Word-Add-in-JavaScript-AddPopulateBindings](https://github.com/OfficeDev/Word-Add-in-JavaScript-AddPopulateBindings), which provides two separate content samples that you can use interchangeably to populate the same binding.</span></span>


## <a name="adding-objects-that-use-additional-office-open-xml-parts"></a><span data-ttu-id="f1c9e-380">添加使用其他 Office Open XML 部分的对象</span><span class="sxs-lookup"><span data-stu-id="f1c9e-380">Adding objects that use additional Office Open XML parts</span></span>


<span data-ttu-id="f1c9e-381">很多内容类型都需要 Office Open XML 数据包中的其他文档部件，这意味着它们要么在另一个部件中引用信息，要么将内容本身存储在一个或多个其他部件中，并在 document.xml 中引用。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-381">Many types of content require additional document parts in the Office Open XML package, meaning that they either reference information in another part or the content itself is stored in one or more additional parts and referenced in document.xml.</span></span>

<span data-ttu-id="f1c9e-382">例如，考虑以下情况：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-382">For example, consider the following:</span></span>


- <span data-ttu-id="f1c9e-383">使用格式样式（如前面图 2 中所示的带样式的文本，以及图 9 中所示的带样式的表格）的内容需要 styles.xml 部件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-383">Content that uses styles for formatting (such as the styled text shown earlier in Figure 2 or the styled table shown in Figure 9) requires the styles.xml part.</span></span>

- <span data-ttu-id="f1c9e-384">图像（如图 3 和图 4 中所示）包括一个（有时是两个）其他部件中的二进制图像数据。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-384">Images (such as those shown in Figures 3 and 4) include the binary image data in one (and sometimes two) additional parts.</span></span>

- <span data-ttu-id="f1c9e-385">SmartArt 图表（如图 10 中所示）需要多个其他部件来说明布局和内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-385">SmartArt diagrams (such as the one shown in Figure 10) require multiple additional parts to describe the layout and content.</span></span>

- <span data-ttu-id="f1c9e-386">图表（如图 11 中所示）需要多个其他部件，包括其自身的关系 (.rels) 部件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-386">Charts (such as the one shown in Figure 11) require multiple additional parts, including their own relationship (.rels) part.</span></span>

<span data-ttu-id="f1c9e-p167">您可以在前面引用的代码示例 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 中看到所有这些内容类型已编辑的标记示例。可以使用前面所述的（以及在引用的代码示例中提供的）同一个 JavaScript 代码插入所有这些内容类型，以在活动选定内容处插入内容，并使用绑定将内容写入到指定的位置。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p167">You can see edited examples of the markup for all of these content types in the previously-referenced code sample [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML). You can insert all of these content types using the same JavaScript code shown earlier (and provided in the referenced code samples) for inserting content at the active selection and writing content to a specified location using bindings.</span></span>

<span data-ttu-id="f1c9e-389">探索示例前，先来看看使用每个内容类型的一些提示。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-389">Before you explore the samples, let's take a look at few tips for working with each of these content types.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="f1c9e-390">请注意，若要保留 document.xml 中引用的其他任何部分，需要保留 document.xml.rels 和要保留的适用部分（如 styles.xml 或图像文件）的关系定义。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-390">Remember, if you are retaining any additional parts referenced in document.xml, you will need to retain document.xml.rels and the relationship definitions for the applicable parts you're keeping, such as styles.xml or an image file.</span></span>


### <a name="working-with-styles"></a><span data-ttu-id="f1c9e-391">使用样式</span><span class="sxs-lookup"><span data-stu-id="f1c9e-391">Working with styles</span></span>

<span data-ttu-id="f1c9e-p168">在使用段落样式或表格样式设置内容格式时，适用与编辑如前所示的直接格式文本示例的标记相同的方法。但是，使用段落样式的标记相当简单，因此在此处作为说明的示例。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p168">The same approach to editing the markup that we looked at for the preceding example with directly-formatted text applies when using paragraph styles or table styles to format your content. However, the markup for working with paragraph styles is considerably simpler, so that is the example described here.</span></span>


#### <a name="editing-the-markup-for-content-using-paragraph-styles"></a><span data-ttu-id="f1c9e-394">使用段落样式编辑内容标记</span><span class="sxs-lookup"><span data-stu-id="f1c9e-394">Editing the markup for content using paragraph styles</span></span>

<span data-ttu-id="f1c9e-395">以下标记表示图 2 中带样式文本示例的正文内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-395">The following markup represents the body content for the styled text example shown in Figure 2.</span></span>


```XML
<w:body>
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Heading1"/>
    </w:pPr>
    <w:r>
      <w:t>This text is formatted using the Heading 1 paragraph style.</w:t>
    </w:r>
  </w:p>
</w:body>
```


> [!NOTE]
> <span data-ttu-id="f1c9e-p169">可以看到，使用样式时，document.xml 中格式化文本的标记非常简单，因为样式包含需要单独引用的所有段落和字体格式。不过，如前所述，建议将样式或直接格式用于不同用途：使用直接格式可以指定文本外观，而不考虑用户文档中的格式；使用段落样式（尤其是内置段落样式名称，如此处所示的“标题 1”），可以让文本格式自动与用户文档进行协调。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p169">As you see, the markup for formatted text in document.xml is considerably simpler when you use a style, because the style contains all of the paragraph and font formatting that you otherwise need to reference individually. However, as explained earlier, you might want to use styles or direct formatting for different purposes: use direct formatting to specify the appearance of your text regardless of the formatting in the user's document; use a paragraph style (particularly a built-in paragraph style name, such as Heading 1 shown here) to have the text formatting automatically coordinate with the user's document.</span></span>

<span data-ttu-id="f1c9e-p170">样式的使用是阅读和了解插入内容标记重要性的一个很好的例子，因为对于此处是否引用另一个文档部件尚不明确。如果此标记中包括样式定义，但不包括 styles.xml 部件，则 document.xml 中的样式信息将会被忽略，不管该样式是否在用户的文档中使用。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p170">Use of a style is a good example of how important it is to read and understand the markup for the content you're inserting, because it's not explicit that another document part is referenced here. If you include the style definition in this markup and don't include the styles.xml part, the style information in document.xml will be ignored regardless of whether or not that style is in use in the user's document.</span></span>

<span data-ttu-id="f1c9e-400">但是，如果查看 styles.xml 部件，您将会看到，在编辑用于您的外接程序的标记时，所需要的仅仅是长段标记中的一小部分：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-400">However, if you take a look at the styles.xml part, you'll see that only a small portion of this long piece of markup is required when editing markup for use in your add-in:</span></span>


- <span data-ttu-id="f1c9e-p171">styles.xml 部件默认包括多个命名空间。如果您仅保留内容必需的样式信息，则在大多数情况下，仅需要保留 **xmlns:w** 命名空间。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p171">The styles.xml part includes several namespaces by default. If you are only retaining the required style information for your content, in most cases you only need to keep the **xmlns:w** namespace.</span></span>

- <span data-ttu-id="f1c9e-403">如果通过外接程序插入标记，并且标记可删除，则样式部件顶部的 **w:docDefaults** 标记内容将被忽略。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-403">The **w:docDefaults** tag content that falls at the top of the styles part will be ignored when your markup is inserted via the add-in and can be removed.</span></span>

- <span data-ttu-id="f1c9e-p172">styles.xml 部件中最长的标记是针对 **w:latentStyles** 标记的，显示在 docDefaults 之后，提供每个可用样式的信息（如“样式”窗格和“样式”库的外观属性）。如果通过外接程序插入内容，并且内容可删除，则此信息也将被忽略。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p172">The largest piece of markup in a styles.xml part is for the **w:latentStyles** tag that appears after docDefaults, which provides information (such as appearance attributes for the Styles pane and Styles gallery) for every available style. This information is also ignored when inserting content via your add-in and so it can be removed.</span></span>

- <span data-ttu-id="f1c9e-p173">在隐藏的样式信息后面，可以看到生成标记的文档中所使用的每个样式的定义。这包括创建新文档时使用的一些可能与内容不相关的默认样式。您可以删除内容未使用的任何样式的定义。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p173">Following the latent styles information, you see a definition for each style in use in the document from which you're markup was generated. This includes some default styles that are in use when you create a new document and may not be relevant to your content. You can delete the definitions for any styles that aren't used by your content.</span></span>


   > [!NOTE]
   > <span data-ttu-id="f1c9e-p174">每个内置标题样式都有关联的字符样式，即相同标题格式的字符样式版本。除非已将标题样式应用为字符样式，否则可以删除它。如果将样式用作字符样式，它显示在 document.xml 中的 run 属性标记 (**w:rPr**)（而不是 paragraph 属性 (**w:pPr**) 标记）内。仅当已将样式应用到部分段落时，才能这么做，但这也会在没有正确应用样式时无意间发生。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p174">Each built-in heading style has an associated Char style that is a character style version of the same heading format. Unless you've applied the heading style as a character style, you can remove it. If the style is used as a character style, it appears in document.xml in a run properties tag ( **w:rPr** ) rather than a paragraph properties ( **w:pPr** ) tag. This should only be the case if you've applied the style to just part of a paragraph, but it can occur inadvertently if the style was incorrectly applied.</span></span>


- <span data-ttu-id="f1c9e-p175">如果在内容中使用内置样式，则无需包括完整的定义，仅需包括样式名称、样式 ID，以及至少一个格式属性，以使强制转换的 Office Open XML 将此样式应用到插入的内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p175">If you're using a built-in style for your content, you don't have to include a full definition. You only must include the style name, style ID, and at least one formatting attribute in order for the coerced Office Open XML to apply the style to your content upon insertion.</span></span>

    <span data-ttu-id="f1c9e-p176">但是，最佳做法是包含一个完整的样式定义（即使它是内置样式的默认值）。如果样式已在目标文档中使用，你的内容将采用该样式的常驻定义，而不考虑 styles.xml 中包含的内容。如果该样式尚未在目标文档中使用，你的内容将使用标记中提供的样式。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p176">However, it's a best practice to include a complete style definition (even if it's the default for built-in styles). If a style is already in use in the destination document, your content will take on the resident definition for the style, regardless of what you include in styles.xml. If the style isn't yet in use in the destination document, your content will use the style definition you provide in the markup.</span></span>

<span data-ttu-id="f1c9e-418">例如，我们需要从图 2 所示的示例文本的 styles.xml 部件保留的唯一内容（使用“Heading 1”样式设置格式）如下。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-418">So, for example, the only content we needed to retain from the styles.xml part for the sample text shown in Figure 2, which is formatted using Heading 1 style, is the following.</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-419">“Heading 1”样式的完整 Word 2013 定义在本示例中已保留。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-419">A complete Word 2013 definition for the Heading 1 style has been retained in this example.</span></span>




```XML
<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
  <pkg:xmlData>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" >
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:next w:val="Normal"/>
        <w:link w:val="Heading1Char"/>
        <w:uiPriority w:val="9"/>
        <w:qFormat/>
        <w:pPr>
          <w:keepNext/>
          <w:keepLines/>
          <w:spacing w:before="240" w:after="0" w:line="259" w:lineRule="auto"/>
          <w:outlineLvl w:val="0"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
          <w:color w:val="2E74B5" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="32"/>
          <w:szCs w:val="32"/>
        </w:rPr>
      </w:style>
    </w:styles>
  </pkg:xmlData>
</pkg:part>
```


#### <a name="editing-the-markup-for-content-using-table-styles"></a><span data-ttu-id="f1c9e-420">使用表格样式编辑内容标记</span><span class="sxs-lookup"><span data-stu-id="f1c9e-420">Editing the markup for content using table styles</span></span>


<span data-ttu-id="f1c9e-p177">当您的内容使用表格样式时，需要与“使用段落样式”中所述的 styles.xml 相关部件相同的部件。也就是说，仅需保留您内容中使用的样式信息，且必须包括名称、ID 以及至少一个格式属性，但如果能包括完整的样式定义以解决所有潜在用户方案将会更好。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p177">When your content uses a table style, you need the same relative part of styles.xml as described for working with paragraph styles. That is, you only need to retain the information for the style you're using in your content, and you must include the name, ID, and at least one formatting attribute, but are better off including a complete style definition to address all potential user scenarios.</span></span>

<span data-ttu-id="f1c9e-423">然而，当查看同时用于 document.xml 中的表格和 styles.xml 中的表格样式定义的标记时，您会看到比使用段落样式时多得多的标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-423">However, when you look at the markup both for your table in document.xml and for your table style definition in styles.xml, you see enormously more markup than when working with paragraph styles.</span></span>


- <span data-ttu-id="f1c9e-p178">在 document.xml 中，即使格式包括在样式内，单元格也会应用该格式。使用表格样式不会减少标记的数量。在内容中使用表格样式的好处是，更新非常简单，且很容易协调多个表格的外观。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p178">In document.xml, formatting is applied by cell even if it's included in a style. Using a table style won't reduce the volume of markup. The benefit of using table styles for the content is for easy updating and easily coordinating the look of multiple tables.</span></span>

- <span data-ttu-id="f1c9e-427">在 styles.xml 中，您将会看到单个表格样式也会有大量标记，这是由于表格样式针对每个表格区域包括多种可能的格式属性类型，如整个表格、标题行、奇数和偶数带状行和列（单独的）、首列等。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-427">In styles.xml, you'll see a substantial amount of markup for a single table style as well, because table styles include several types of possible formatting attributes for each of several table areas, such as the entire table, heading rows, odd and even banded rows and columns (separately), the first column, etc.</span></span>


### <a name="working-with-images"></a><span data-ttu-id="f1c9e-428">使用图像</span><span class="sxs-lookup"><span data-stu-id="f1c9e-428">Working with images</span></span>


<span data-ttu-id="f1c9e-p179">图像的标记包括一个对至少一个部件的引用，该部件包含用以说明图像的二进制数据。对于复杂的图像，可能有数百页的标记，并且无法进行编辑。由于无需涉及二进制部件，在使用结构化编辑器（如 Visual Studio）时可以简单地将其折叠，因此您仍可以很轻松地查看并编辑数据包的其余部分。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p179">The markup for an image includes a reference to at least one part that includes the binary data to describe your image. For a complex image, this can be hundreds of pages of markup and you can't edit it. Since you don't ever have to touch the binary part(s), you can simply collapse it if you're using a structured editor such as Visual Studio, so that you can still easily review and edit the rest of the package.</span></span>

<span data-ttu-id="f1c9e-p180">如果查看图 3 中所示简单图像的示例标记（该标记可用于前面引用的示例代码 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 中），您会看到 document.xml 中的图像标记包括大小和位置信息，以及对包含二进制图像数据的部件的关系引用。该引用包括在 **a:blip** 标记中，如下所示：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p180">If you check out the example markup for the simple image shown earlier in Figure 3, available in the previously-referenced code sample [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML), you'll see that the markup for the image in document.xml includes size and position information as well as a relationship reference to the part that contains the binary image data. That reference is included in the **a:blip** tag, as follows:</span></span>




```XML
<a:blip r:embed="rId4" cstate="print">
```

<span data-ttu-id="f1c9e-p181">请注意，由于关系引用由 (**r:embed="rID4"**) 明确使用，并且为了呈现图像，相关部件是必需的，如果 Office Open XML 数据包中未包括二进制数据，则会出现错误。这与前面所述的 styles.xml 有所不同，在 styles.xml 中不会引发错误，因为没有明确引用关系，且关系是为内容提供属性的部件，而非其本身成为内容的一部分。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p181">Be aware that, because a relationship reference is explicitly used ( **r:embed="rID4"** ) and that related part is required in order to render the image, if you don't include the binary data in your Office Open XML package, you will get an error. This is different from styles.xml, explained previously, which won't throw an error if omitted since the relationship is not explicitly referenced and the relationship is to a part that provides attributes to the content (formatting) rather than being part of the content itself.</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-p182">查看标记时，请注意 a:blip 标记中使用的其他命名空间。在 document.xml 中，**xlmns:a** 命名空间（主 drawingML 命名空间）被动态置于使用 drawingML 引用的开始部分，而非 document.xml 部分的顶部。然而，关系命名空间 (r) 必须按原样保留在 document.xml 开头位置。请检查图片标记是否有其他命名空间要求。请注意，无需记住哪种内容类型需要哪个命名空间，通过查看整个 document.xml 中的标记前缀，就能很容易地分辨出来。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p182">When you review the markup, notice the additional namespaces used in the a:blip tag. You'll see in document.xml that the  **xlmns:a** namespace (the main drawingML namespace) is dynamically placed at the beginning of the use of drawingML references rather than at the top of the document.xml part. However, the relationships namespace (r) must be retained where it appears at the start of document.xml. Check your picture markup for additional namespace requirements. Remember that you don't have to memorize which types of content require what namespaces, you can easily tell by reviewing the prefixes of the tags throughout document.xml.</span></span>


### <a name="understanding-additional-image-parts-and-formatting"></a><span data-ttu-id="f1c9e-441">了解其他图像部分和格式</span><span class="sxs-lookup"><span data-stu-id="f1c9e-441">Understanding additional image parts and formatting</span></span>


<span data-ttu-id="f1c9e-p183">在图像上使用某些 Office 图片格式效果时（如图 4 中所示的图像，该图像除使用图片样式之外，还使用已调整的亮度和对比度设置），可能需要针对图像数据的 HD 格式副本的第二个二进制数据部件。考虑分层效果的格式需要这个额外的 HD 格式，并且对该格式的引用显示在 document.xml 中，类似于以下内容：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p183">When you use some Office picture formatting effects on your image, such as for the image shown in Figure 4, which uses adjusted brightness and contrast settings (in addition to picture styling), a second binary data part for an HD format copy of the image data may be required. This additional HD format is required for formatting considered a layering effect, and the reference to it appears in document.xml similar to the following:</span></span>


```XML
<a14:imgLayer r:embed="rId5">
```

<span data-ttu-id="f1c9e-444">请在 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 代码示例中参阅图 4 中所示的（使用分层效果等）带格式图像所需的标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-444">See the required markup for the formatted image shown in Figure 4 (which uses layering effects among others) in the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample.</span></span>


### <a name="working-with-smartart-diagrams"></a><span data-ttu-id="f1c9e-445">使用 SmartArt 图表</span><span class="sxs-lookup"><span data-stu-id="f1c9e-445">Working with SmartArt diagrams</span></span>


<span data-ttu-id="f1c9e-p184">SmartArt 图表具有四个关联的部件，但始终需要的只有两个。您可以检查 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 代码示例中的 SmartArt 标记示例。首先，了解一下每个部件的简要说明，以及为什么需要/不需要这些部件：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p184">A SmartArt diagram has four associated parts, but only two are always required. You can examine an example of SmartArt markup in the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample. First, take a look at a brief description of each of the parts and why they are or are not required:</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-449">如果内容包括多个图表，它们会进行连续编号，替换此处列出的文件名中的“1”。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-449">If your content includes more than one diagram, they will be numbered consecutively, replacing the 1 in the file names listed here.</span></span>


- <span data-ttu-id="f1c9e-p185">layout1.xml：此部件是必需的。它包括布局外观和功能的标记定义。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p185">layout1.xml: This part is required. It includes the markup definition for the layout appearance and functionality.</span></span>

- <span data-ttu-id="f1c9e-p186">data1.xml：此部件是必需的。它包括图表实例中使用的数据。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p186">data1.xml: This part is required. It includes the data in use in your instance of the diagram.</span></span>

- <span data-ttu-id="f1c9e-454">drawing1.xml：此部件不是始终必需的，但如果将自定义格式应用到图表实例中的元素（如直接格式化各个形状），则可能需要保留它。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-454">drawing1.xml: This part is not always required but if you apply custom formatting to elements in your instance of a diagram, such as directly formatting individual shapes, you might need to retain it.</span></span>

- <span data-ttu-id="f1c9e-p187">colors1.xml：此部件不是必需的。它包括颜色样式信息，但图表的颜色会默认与目标文档中活动格式主题的颜色协调，这取决于保存 Office Open XML 标记之前，从 Word 中的“SmartArt 工具设计”选项卡应用的 SmartArt 颜色样式。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p187">colors1.xml: This part is not required. It includes color style information, but the colors of your diagram will coordinate by default with the colors of the active formatting theme in the destination document, based on the SmartArt color style you apply from the SmartArt Tools design tab in Word before saving out your Office Open XML markup.</span></span>

- <span data-ttu-id="f1c9e-p188">quickStyles1.xml：此部件不是必需的。与颜色部件相似，如果图表将采用已应用 SmartArt 样式（可用于目标文档中）的定义（也就是说，它会自动与目标文档中的格式主题协调），则可以删除此部件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p188">quickStyles1.xml: This part is not required. Similar to the colors part, you can remove this as your diagram will take on the definition of the applied SmartArt style that's available in the destination document (that is, it will automatically coordinate with the formatting theme in the destination document).</span></span>


> [!TIP]
> <span data-ttu-id="f1c9e-p189">SmartArt layout1.xml 文件很好地示范了可以进一步修整标记的位置，但额外花时间这样做可能并不值得，因为只会删除与整个包相关的少量标记。若要从标记中清除所有代码行，可以删除 **dgm:sampData** 标记及其内容。此示例数据定义了如何在 SmartArt 样式库中显示图表的缩略图预览。不过，如果省略，使用的是默认示例数据。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p189">The SmartArt layout1.xml file is a good example of places you may be able to further trim your markup but might not be worth the extra time to do so (because it removes such a small amount of markup relative to the entire package). If you would like to get rid of every last line you can of markup, you can delete the **dgm:sampData** tag and its contents. This sample data defines how the thumbnail preview for the diagram will appear in the SmartArt styles galleries. However, if it's omitted, default sample data is used.</span></span>

<span data-ttu-id="f1c9e-p190">请注意，document.xml 中 SmartArt 图表的标记包含对布局、数据、颜色和快速样式部件的关系 ID 引用。在删除这些部件及其关系定义（由于删除的是这些关系，则确定是此操作的最佳做法）时，可以删除 document.xml 中对颜色和样式部件的引用，但如果保留它们，则不会产生错误，因为它们不是将图表插入文档中所必需的。在 **dgm:relIds** 标记中的 document.xml 中查找这些引用。不管是否执行此步骤，都要保留所需的布局和数据部件的关系 ID 引用。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p190">Be aware that the markup for a SmartArt diagram in document.xml contains relationship ID references to the layout, data, colors, and quick styles parts. You can delete the references in document.xml to the colors and styles parts when you delete those parts and their relationship definitions (and it's certainly a best practice to do so, since you're deleting those relationships), but you won't get an error if you leave them, since they aren't required for your diagram to be inserted into a document. Find these references in document.xml in the  **dgm:relIds** tag. Regardless of whether or not you take this step, retain the relationship ID references for the required layout and data parts.</span></span>


### <a name="working-with-charts"></a><span data-ttu-id="f1c9e-467">使用图表</span><span class="sxs-lookup"><span data-stu-id="f1c9e-467">Working with charts</span></span>


<span data-ttu-id="f1c9e-p191">类似于 SmartArt 图表，图表包含多个其他部件。但是，图表的配置与 SmartArt 有所不同，区别在于图表有其自身的关系文件。以下是图表所需的且可删除的文档部件说明：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p191">Similar to SmartArt diagrams, charts contain several additional parts. However, the setup for charts is a bit different from SmartArt, in that a chart has its own relationship file. Following is a description of required and removable document parts for a chart:</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-471">对于 SmartArt 图表，如果内容包括多个图表，则会将它们连续编号，替换此处列出的文件名称中的“1”。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-471">As with SmartArt diagrams, if your content includes more than one chart, they will be numbered consecutively, replacing the 1 in the file names listed here.</span></span>


- <span data-ttu-id="f1c9e-472">document.xml.rels 引用了包含图表 (chart1.xml) 描述数据的必需部分。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-472">In document.xml.rels, you'll see a reference to the required part that contains the data that describes the chart (chart1.xml).</span></span>

- <span data-ttu-id="f1c9e-473">您还会看到 Office Open XML 数据包中每个图表单独的关系文件，如 chart1.xml.rels。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-473">You also see a separate relationship file for each chart in your Office Open XML package, such as chart1.xml.rels.</span></span>

    <span data-ttu-id="f1c9e-p192">chart1.xml.rels 中共引用了三个文件，但只有一个是必需的。其中包括二进制 Excel 工作簿数据（必需）和可以删除的颜色和样式部件（colors1.xml 和 styles1.xml）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p192">There are three files referenced in chart1.xml.rels, but only one is required. These include the binary Excel workbook data (required) and the color and style parts (colors1.xml and styles1.xml) that you can remove.</span></span>

<span data-ttu-id="f1c9e-p193">可以在本机 Word 2013 中创建并编辑的图表为 Excel 2013 图表，其数据在作为二进制数据嵌入 Office Open XML 数据包的 Excel 工作簿上进行维护。与图像的二进制数据部件类似，此 Excel 二进制数据也是必需的，但此部件中没有要编辑的内容。因此您只需在编辑器中折叠此部件，从而避免需要手动滚动全部内容来检查 Office Open XML 数据包的剩余部分。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p193">Charts that you can create and edit natively in Word 2013 are Excel 2013 charts, and their data is maintained on an Excel worksheet that's embedded as binary data in your Office Open XML package. Like the binary data parts for images, this Excel binary data is required, but there's nothing to edit in this part. So you can just collapse the part in the editor to avoid having to manually scroll through it all to examine the rest of your Office Open XML package.</span></span>

<span data-ttu-id="f1c9e-p194">但是，类似于 SmartArt，您可以删除颜色和样式部件。如果使用了可用的图表样式和颜色样式来为图表设置格式，则图表将在插入目标文档时自动呈现为适用的格式。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p194">However, similar to SmartArt, you can delete the colors and styles parts. If you've used the chart styles and color styles available in to format your chart, the chart will take on the applicable formatting automatically when it is inserted into the destination document.</span></span>

<span data-ttu-id="f1c9e-481">请在 [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) 代码示例中参阅图 11 中所示的示例图表的已编辑标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-481">See the edited markup for the example chart shown in Figure 11 in the [Word-Add-in-Load-and-write-Open-XML](https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML) code sample.</span></span>


## <a name="editing-the-office-open-xml-for-use-in-your-task-pane-add-in"></a><span data-ttu-id="f1c9e-482">编辑 Office Open XML 以用于任务窗格外接程序</span><span class="sxs-lookup"><span data-stu-id="f1c9e-482">Editing the Office Open XML for use in your task pane add-in</span></span>


<span data-ttu-id="f1c9e-p195">您已经了解如何标识并编辑标记中的内容。如果在查看文档生成的大量 Office Open XML 数据包时仍觉得任务似乎有困难，以下是推荐步骤的快速摘要，可帮助您快速编辑数据包：</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p195">You've already seen how to identify and edit the content in your markup. If the task still seems difficult when you take a look at the massive Office Open XML package generated for your document, following is a quick summary of recommended steps to help you edit that package down quickly:</span></span>


> [!NOTE]
> <span data-ttu-id="f1c9e-485">请记住，您可以使用数据包中的所有 .rels 部件作为地图，以快速检查可以删除的文档部件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-485">Remember that you can use all .rels parts in the package as a map to quickly check for document parts that you can remove.</span></span>


1. <span data-ttu-id="f1c9e-p196">在 Visual Studio 2015 中打开平展的 XML 文件，并按 Ctrl+K 和 Ctrl+D 设置文件格式。然后使用左侧的折叠/展开按钮折叠需要删除的部件。您可能还想要折叠需要但无需编辑的长部件（如图像文件的 base64 二进制数据），以使标记可以更快速更容易地进行可视化浏览。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p196">Open the flattened XML file in Visual Studio 2015 and press Ctrl+K, Ctrl+D to format the file. Then use the collapse/expand buttons on the left to collapse the parts you know you need to remove. You might also want to collapse long parts you need, but know you won't need to edit (such as the base64 binary data for an image file), making the markup faster and easier to visually scan.</span></span>

2. <span data-ttu-id="f1c9e-489">在准备用于加载项的 Office Open XML 标记时，文档包的几个部分几乎总是可以删除。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-489">There are several parts of the document package that you can almost always remove when you are preparing Office Open XML markup for use in your add-in.</span></span> <span data-ttu-id="f1c9e-490">建议首先删除这些部分（及其关联的关系定义），这将立即大大减少包。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-490">You might want to start by removing these (and their associated relationship definitions), which will greatly reduce the package right away.</span></span> <span data-ttu-id="f1c9e-491">这些包括 theme1、fontTable、设置、webSettings、缩略图以及核心和加载项属性文件以及任何 `taskpane` 或 `webExtension` 部分。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-491">These include the theme1, fontTable, settings, webSettings, thumbnail, both the core and add-in properties files, and any `taskpane` or `webExtension` parts.</span></span>

3. <span data-ttu-id="f1c9e-p198">删除与您内容不相关的任何部件，例如，不需要的脚注、页眉或页脚。请记住，还要删除其关联的关系。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p198">Remove any parts that don't relate to your content, such as footnotes, headers, or footers that you don't require. Again, remember to also delete their associated relationships.</span></span>

4. <span data-ttu-id="f1c9e-p199">查看 document.xml.rels 部件以查看该部件中引用的任何文件（如图像文件、样式部件或 SmartArt 图表部件）是否是您的内容所必需的。删除您的内容不需要的所有部件的关系，并确认还删除了其关联的部件。如果您的内容不需要 document.xml.rels 中引用的任何文档部件，则也可以删除该文件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p199">Review the document.xml.rels part to see if any files referenced in that part are required for your content, such as an image file, the styles part, or SmartArt diagram parts. Delete the relationships for any parts your content doesn't require and confirm that you have also deleted the associated part. If your content doesn't require any of the document parts referenced in document.xml.rels, you can delete that file also.</span></span>

5. <span data-ttu-id="f1c9e-497">如果您的内容具有其他 .rels 部件（如 chart#.xml.rels），则查看是否有可以删除的其他引用部件（如图表的快速样式），并从该文件中和关联的部件中删除关系。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-497">If your content has an additional .rels part (such as chart#.xml.rels), review it to see if there are other parts referenced there that you can remove (such as quick styles for charts) and delete both the relationship from that file as well as the associated part.</span></span>

6. <span data-ttu-id="f1c9e-p200">编辑 document.xml 以删除部件中未引用的命名空间、删除内容未包括分节符时的节属性，以及删除与您想要插入的内容不相关的所有标记。如果插入的是形状或文本框，您可能还想删除扩展的回退标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p200">Edit document.xml to remove namespaces not referenced in the part, section properties if your content doesn't include a section break, and any markup that's not related to the content that you want to insert. If inserting shapes or text boxes, you might also want to remove extensive fallback markup.</span></span>

7. <span data-ttu-id="f1c9e-500">对删除大量标记不会影响您内容的情况下，对任何其他所需部件进行编辑，如样式部件。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-500">Edit any additional required parts where you know that you can remove substantial markup without affecting your content, such as the styles part.</span></span>

<span data-ttu-id="f1c9e-p201">执行了前面七个步骤之后，您就有可能剪切 90% - 100% 的可删除标记，这取决于您的内容。大多数情况下，可以按照您想要剪裁的多少进行。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p201">After you've taken the preceding seven steps, you've likely cut between about 90 and 100 percent of the markup you can remove, depending on your content. In most cases, this is likely to be as far as you want to trim.</span></span>

<span data-ttu-id="f1c9e-503">无论是保留，还是选择深入内容以查找可以剪切的每一行标记，都请记住，您可以使用前面引用的代码示例 [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) 作为便签簿以快速简单地测试已编辑标记。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-503">Regardless of whether you leave it here or choose to delve further into your content to find every last line of markup you can cut, remember that you can use the previously-referenced code sample [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML) as a scratch pad to quickly and easily test your edited markup.</span></span>


> [!TIP]
> <span data-ttu-id="f1c9e-p202">如果在开发期间更新现有解决方案中的 Office Open XML 代码片段，请先清除 Internet 临时文件，再重新运行解决方案，以更新代码使用的 Office Open XML。解决方案中 XML 文件包含的标记会缓存到计算机。当然，可以从默认 Web 浏览器中清除 Internet 临时文件。若要在 Visual Studio 2015 中访问 Internet 选项并删除这些设置，请选择“调试”\*\*\*\* 菜单中的“选项和设置”\*\*\*\*。然后在“环境”\*\*\*\* 下，依次选择“Web 浏览器”\*\*\*\* 和“Internet Explorer 选项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p202">If you update an Office Open XML snippet in an existing solution while developing, clear temporary Internet files before you run the solution again to update the Office Open XML used by your code. Markup that's included in your solution in XML files is cached on your computer. You can, of course, clear temporary Internet files from your default web browser. To access Internet options and delete these settings from inside Visual Studio 2015, on the  **Debug** menu, choose **Options and Settings**. Then, under  **Environment**, choose  **Web Browser** and then choose **Internet Explorer Options**.</span></span>


## <a name="creating-an-add-in-for-both-template-and-stand-alone-use"></a><span data-ttu-id="f1c9e-509">创建用于模板和独立使用的加载项</span><span class="sxs-lookup"><span data-stu-id="f1c9e-509">Creating an add-in for both template and stand-alone use</span></span>


<span data-ttu-id="f1c9e-p203">在本主题中，您了解到外接程序中可以使用 Office Open XML 进行操作的多个示例。我们了解了可以使用 Office Open XML 强制转换类型插入到文档中的各种多种格式的内容类型示例，以及在选定内容或指定（限制）位置插入该内容的 JavaScript 方法。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p203">In this topic, you've seen several examples of what you can do with Office Open XML in your add-ins for . We've looked at a wide range of rich content type examples that you can insert into documents by using the Office Open XML coercion type, together with the JavaScript methods for inserting that content at the selection or to a specified (bound) location.</span></span>

<span data-ttu-id="f1c9e-p204">如果您创建的是可独立使用（即从应用商店或专有服务器位置插入的），也可在预先创建的模板（设计为与外接程序一起使用）中使用的外接程序，您还需要了解什么内容？答案应该是，您已经了解了所有所需的内容。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p204">So, what else do you need to know if you're creating your add-in both for stand-alone use (that is, inserted from the Store or a proprietary server location) and for use in a pre-created template that's designed to work with your add-in? The answer might be that you already know all you need.</span></span>

<span data-ttu-id="f1c9e-p205">无论外接程序是设计为独立使用，还是与模板一起使用，给定内容类型和插入方法的标记都相同。如果您使用的模板是设计为与外接程序一起使用，请确保 JavaScript 包括回退，该回退用于说明引用的内容可能存在于文档中的方案（如“[添加并绑定到命名内容控件](#add-and-bind-to-a-named-content-control)”一节中所示的绑定示例中所演示的）。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p205">The markup for a given content type and methods for inserting it are the same whether your add-in is designed to stand-alone or work with a template. If you are using templates designed to work with your add-in, just be sure that your JavaScript includes callbacks that account for scenarios where referenced content might already exist in the document (such as demonstrated in the binding example shown in the section [Add and bind to a named content control](#add-and-bind-to-a-named-content-control)).</span></span>

<span data-ttu-id="f1c9e-p206">通过应用使用模板时，无论外接程序是在用户创建文档时常驻在模板中，还是外接程序将插入模板，您都可能还想结合 API 的其他元素，以帮助您创建更可靠的交互式体验。例如，您可能想要在自定义 XML 部件中包括标识数据，以便可以使用此标识数据确定模板类型，从而为用户提供特定于模板的选项。若要了解有关如何在外接程序中使用自定义 XML 的详细信息，请参阅下面的“其他资源”部分。</span><span class="sxs-lookup"><span data-stu-id="f1c9e-p206">When using templates with your app, whether the add-in will be resident in the template at the time that the user created the document or the add-in will be inserting a template, you might also want to incorporate other elements of the API to help you create a more robust, interactive experience. For example, you may want to include identifying data in a customXML part that you can use to determine the template type in order to provide template-specific options to the user. To learn more about how to work with custom XML in your add-ins, see the additional resources that follow.</span></span>


## <a name="see-also"></a><span data-ttu-id="f1c9e-519">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f1c9e-519">See also</span></span>

- [<span data-ttu-id="f1c9e-520">适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="f1c9e-520">JavaScript API for Office </span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)
- <span data-ttu-id="f1c9e-521">[标准 ECMA-376：Office Open XML 文件格式](https://www.ecma-international.org/publications/standards/Ecma-376.htm)（其中收录了 Open XML 的完整语言参考和相关文档）</span><span class="sxs-lookup"><span data-stu-id="f1c9e-521">[Standard ECMA-376: Office Open XML File Formats](https://www.ecma-international.org/publications/standards/Ecma-376.htm) (access the complete language reference and related documentation on Open XML here)</span></span>
- [<span data-ttu-id="f1c9e-522">探索适用于 Office 的 JavaScript API：数据绑定和自定义 XML 部分</span><span class="sxs-lookup"><span data-stu-id="f1c9e-522">Exploring the JavaScript API for Office: Data Binding and Custom XML Parts</span></span>](https://msdn.microsoft.com/magazine/dn166930.aspx)
