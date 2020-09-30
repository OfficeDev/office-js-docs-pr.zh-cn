---
title: 从网页中打开 Excel 并嵌入 Office 外接程序
description: 从网页中打开 Excel 并嵌入 Office 外接程序。
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: 00846ca5ca05e65fd75629f5aad0e4fb3d947ab1
ms.sourcegitcommit: 42202d7e2ac24dffa77cf937f5697a1cd79ee790
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2020
ms.locfileid: "48308542"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a><span data-ttu-id="e59d5-103">从网页中打开 Excel 并嵌入 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="e59d5-103">Open Excel from your web page and embed your Office Add-in</span></span>

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="网页上的 Excel 按钮的图像使用外接程序打开新的 Excel 文档，并将其置于嵌入式和自动打开状态。":::

<span data-ttu-id="e59d5-105">扩展 SaaS web 应用程序，以便客户可以直接将其数据从网页中直接打开到 Microsoft Excel。</span><span class="sxs-lookup"><span data-stu-id="e59d5-105">Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel.</span></span> <span data-ttu-id="e59d5-106">一个常见的情况是，客户将使用 web 应用程序中的数据。</span><span class="sxs-lookup"><span data-stu-id="e59d5-106">A common scenario is that customers will be working with data in your web application.</span></span> <span data-ttu-id="e59d5-107">然后，他们需要将数据复制到 Excel 文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-107">Then they’ll want to copy the data into an Excel document.</span></span> <span data-ttu-id="e59d5-108">例如，他们可能想要使用 Excel 执行其他分析。</span><span class="sxs-lookup"><span data-stu-id="e59d5-108">For example, they may want to perform additional analysis using Excel.</span></span> <span data-ttu-id="e59d5-109">通常情况下，客户需要将数据导出到文件（如 .csv 文件），然后将该数据导入到 Excel 中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-109">Typically, the customer is required to export the data to a file, such as a .csv file, and then import that data into Excel.</span></span> <span data-ttu-id="e59d5-110">他们还必须手动将 Office 加载项添加到文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-110">They also have to manually add your Office Add-in to the document.</span></span>

<span data-ttu-id="e59d5-111">将步骤数减少为单个按钮单击可生成并打开 Excel 文档的网页。</span><span class="sxs-lookup"><span data-stu-id="e59d5-111">Reduce the number of steps to a single button click on your web page that generates and opens the Excel document.</span></span> <span data-ttu-id="e59d5-112">您还可以将 Office 外接程序嵌入文档中并在文档打开时显示它。</span><span class="sxs-lookup"><span data-stu-id="e59d5-112">You can also embed your Office Add-in inside the document and display it when the document opens.</span></span> <span data-ttu-id="e59d5-113">这可确保客户仍有权访问应用程序功能。</span><span class="sxs-lookup"><span data-stu-id="e59d5-113">This ensures the customer still has access to your application features.</span></span> <span data-ttu-id="e59d5-114">当文档打开时，客户选择的数据和你的 Office 外接程序已可供他们继续工作。</span><span class="sxs-lookup"><span data-stu-id="e59d5-114">When the document opens, the data the customer selected, and your Office Add-in is already available for them to continue working.</span></span>

<span data-ttu-id="e59d5-115">本文介绍在您自己的 SaaS web 应用程序中实现此方案的代码和技术。</span><span class="sxs-lookup"><span data-stu-id="e59d5-115">This article shows you code and techniques for implementing this scenario in your own SaaS web application.</span></span>

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a><span data-ttu-id="e59d5-116">创建新的 Excel 文档并嵌入 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="e59d5-116">Create a new Excel document and embed an Office Add-in</span></span>

<span data-ttu-id="e59d5-117">首先，我们来了解如何从网页创建 Excel 文档，并将外接程序嵌入文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-117">First, let’s learn how to create an Excel document from a web page, and embed an add-in into the document.</span></span> <span data-ttu-id="e59d5-118">[OFFICE OOXML 嵌入加载项代码示例](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)显示了如何将[脚本实验室加载项](https://appsource.microsoft.com/product/office/wa104380862)嵌入到新的 Office 文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-118">The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span></span> <span data-ttu-id="e59d5-119">虽然本示例适用于任何 Office 文档，但我们只是在本文中重点介绍 Excel 电子表格。</span><span class="sxs-lookup"><span data-stu-id="e59d5-119">Although the sample works with any Office document, we’ll just focus on Excel spreadsheets in this article.</span></span> <span data-ttu-id="e59d5-120">使用以下步骤生成并运行示例。</span><span class="sxs-lookup"><span data-stu-id="e59d5-120">Use the following steps to build and run the sample.</span></span>

1. <span data-ttu-id="e59d5-121">将示例代码从  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip 您的计算机上的文件夹中提取出来。</span><span class="sxs-lookup"><span data-stu-id="e59d5-121">Extract the sample code from  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip into a folder on your computer.</span></span>
2. <span data-ttu-id="e59d5-122">若要生成并运行示例，请按照自述文件的 " **使用项目"** 部分中的步骤操作。</span><span class="sxs-lookup"><span data-stu-id="e59d5-122">To build and run the sample, follow the steps in the **To use the project** section of the readme.</span></span>
3. <span data-ttu-id="e59d5-123">运行示例时，它将显示一个类似于以下屏幕截图的网页。</span><span class="sxs-lookup"><span data-stu-id="e59d5-123">When you run the sample it will display a web page similar to the following screen shot.</span></span> <span data-ttu-id="e59d5-124">使用网页创建一个在打开时包含脚本实验室的新 Excel 文档。</span><span class="sxs-lookup"><span data-stu-id="e59d5-124">Use the web page to create a new Excel document that contains Script Lab when it opens.</span></span>
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="网页上的 Excel 按钮的图像使用外接程序打开新的 Excel 文档，并将其置于嵌入式和自动打开状态。":::

### <a name="how-the-sample-works"></a><span data-ttu-id="e59d5-126">示例的工作原理</span><span class="sxs-lookup"><span data-stu-id="e59d5-126">How the sample works</span></span>

<span data-ttu-id="e59d5-127">示例代码使用 OOXML SDK 将脚本实验室外接程序嵌入到您选择的 Excel 文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-127">The sample code uses the OOXML SDK to embed the Script Lab add-in to the Excel document that you choose.</span></span> <span data-ttu-id="e59d5-128">以下信息取自自述文件中的 " [**代码"** 部分](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) 。</span><span class="sxs-lookup"><span data-stu-id="e59d5-128">The following Information is taken from the [**About the code** section](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) in the readme file.</span></span>

<span data-ttu-id="e59d5-129">文件 **Home.aspx.cs**：</span><span class="sxs-lookup"><span data-stu-id="e59d5-129">The file **Home.aspx.cs**:</span></span>

- <span data-ttu-id="e59d5-130">提供按钮事件处理程序和基本 UI 操作。</span><span class="sxs-lookup"><span data-stu-id="e59d5-130">Provides the button event handlers and basic UI manipulation.</span></span>
- <span data-ttu-id="e59d5-131">使用标准 ASP.NET 技术上载和下载文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-131">Uses standard ASP.NET techniques to upload and download the file.</span></span>
- <span data-ttu-id="e59d5-132">使用上传的文件的文件扩展名 (.xlsx、.docx 或 .pptx) 来确定文件的类型。</span><span class="sxs-lookup"><span data-stu-id="e59d5-132">Uses the file name extension of the uploaded file (xlsx, docx, or pptx) to determine the type of file.</span></span> <span data-ttu-id="e59d5-133">需要在开始时执行此操作，因为 Open XML SDK 通常对每种类型的文件都具有不同的 Api。</span><span class="sxs-lookup"><span data-stu-id="e59d5-133">This needs to be done at the outset because the Open XML SDK generally has distinct APIs for each type of file.</span></span>
- <span data-ttu-id="e59d5-134">调用 **OOXMLHelper** 以验证文件并调用 **AddInEmbedder** 以在文件中嵌入脚本实验室并将其设置为自动打开。</span><span class="sxs-lookup"><span data-stu-id="e59d5-134">Calls into the **OOXMLHelper** to validate the file and calls into the **AddInEmbedder** to embed Script Lab in the file and set to automatically open.</span></span>

<span data-ttu-id="e59d5-135">文件 **AddInEmbedder.cs**：</span><span class="sxs-lookup"><span data-stu-id="e59d5-135">The file **AddInEmbedder.cs**:</span></span>

- <span data-ttu-id="e59d5-136">提供主要业务逻辑，在此示例中，是一种嵌入脚本实验室的方法。</span><span class="sxs-lookup"><span data-stu-id="e59d5-136">Provides the main business logic, which in this sample is a method that embeds Script Lab.</span></span>
- <span data-ttu-id="e59d5-137">根据文件的类型，对 OOXML 帮助程序进行调用。</span><span class="sxs-lookup"><span data-stu-id="e59d5-137">Makes calls into the OOXML helper based on the type of the file.</span></span>

<span data-ttu-id="e59d5-138">文件 **OOXMLHelper.cs**：</span><span class="sxs-lookup"><span data-stu-id="e59d5-138">The file **OOXMLHelper.cs**:</span></span>

- <span data-ttu-id="e59d5-139">提供所有详细的 OOXML 操作。</span><span class="sxs-lookup"><span data-stu-id="e59d5-139">Provides all the detailed OOXML manipulation.</span></span>
- <span data-ttu-id="e59d5-140">使用用于验证 Office 文件的标准技术，这只是调用 **文档的 Open** 方法。</span><span class="sxs-lookup"><span data-stu-id="e59d5-140">Uses a standard technique for validating the Office file, which is simply to call the **Document.Open** method on it.</span></span> <span data-ttu-id="e59d5-141">如果文件无效，则该方法将引发异常。</span><span class="sxs-lookup"><span data-stu-id="e59d5-141">If the file is invalid, the method throws an exception.</span></span>
- <span data-ttu-id="e59d5-142">主要包含由 Open XML 2.5 SDK 生产力工具生成的代码，这些工具在 [OPEN xml 2.5 sdk](/office/open-xml/open-xml-sdk)的链接中可用。</span><span class="sxs-lookup"><span data-stu-id="e59d5-142">Contains mainly code that was generated by the Open XML 2.5 SDK Productivity Tools which are available at the link for the [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk).</span></span>

<span data-ttu-id="e59d5-143">**OOXMLHelper.cs**文件中的**GenerateWebExtensionPart1Content**方法将引用设置为 Microsoft AppSource 中的脚本实验室的 ID：</span><span class="sxs-lookup"><span data-stu-id="e59d5-143">The **GenerateWebExtensionPart1Content** method in the **OOXMLHelper.cs** file sets the reference to the ID of Script Lab in Microsoft AppSource:</span></span>

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- <span data-ttu-id="e59d5-144">**StoreType**值为 "OMEX"，为 Microsoft AppSource 的别名。</span><span class="sxs-lookup"><span data-stu-id="e59d5-144">The **StoreType** value is "OMEX", an alias for Microsoft AppSource.</span></span>
- <span data-ttu-id="e59d5-145">**存储**值为脚本实验室的 Microsoft AppSource 区域性部分中的 "en-us"。</span><span class="sxs-lookup"><span data-stu-id="e59d5-145">The **Store** value is "en-US" found in the Microsoft AppSource culture section for Script Lab.</span></span>
- <span data-ttu-id="e59d5-146">**Id**值是脚本实验室的 Microsoft APPSOURCE 资产 Id。</span><span class="sxs-lookup"><span data-stu-id="e59d5-146">The **Id** value is the Microsoft AppSource asset ID for Script Lab.</span></span>

<span data-ttu-id="e59d5-147">如果要从文件共享目录中设置自动打开的外接程序，您将使用不同的值：</span><span class="sxs-lookup"><span data-stu-id="e59d5-147">If you are setting up an add-in from a file share catalog for auto-open, you will use different values:</span></span>

<span data-ttu-id="e59d5-148">**StoreType**值为 "FileSystem"。</span><span class="sxs-lookup"><span data-stu-id="e59d5-148">The **StoreType** value is "FileSystem".</span></span>

- <span data-ttu-id="e59d5-149">**存储**值是网络共享的 URL;例如，" \\ \\ MyComputer \\ MySharedFolder"。</span><span class="sxs-lookup"><span data-stu-id="e59d5-149">The **Store** value is the URL of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span> <span data-ttu-id="e59d5-150">这应该是在 Office 信任中心中显示为共享的受信任目录地址的确切 URL。</span><span class="sxs-lookup"><span data-stu-id="e59d5-150">This should be the exact URL that appears as the share's Trusted Catalog Address in the Office Trust Center.</span></span>
- <span data-ttu-id="e59d5-151">**Id**值是加载项清单中的应用程序 Id。</span><span class="sxs-lookup"><span data-stu-id="e59d5-151">The **Id** value is the app ID in the add-ins manifest.</span></span>
> [!NOTE]
> <span data-ttu-id="e59d5-152">有关这些属性的可选值的详细信息，请参阅 [自动打开包含文档的任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)。</span><span class="sxs-lookup"><span data-stu-id="e59d5-152">For more information about alternative values for these attributes, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).</span></span>

## <a name="use-the-fluent-ui"></a><span data-ttu-id="e59d5-153">使用熟知的 UI</span><span class="sxs-lookup"><span data-stu-id="e59d5-153">Use the Fluent UI</span></span>

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="网页上的 Excel 按钮的图像使用外接程序打开新的 Excel 文档，并将其置于嵌入式和自动打开状态。":::

<span data-ttu-id="e59d5-155">最佳做法是使用熟知的 UI 来帮助用户在 Microsoft 产品之间进行转换。</span><span class="sxs-lookup"><span data-stu-id="e59d5-155">A best practice is to use the Fluent UI to help your users transition between Microsoft products.</span></span> <span data-ttu-id="e59d5-156">应始终使用 Office 图标来指示将从您的网页启动哪个 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="e59d5-156">You should always use an Office icon to indicate which Office application will be launched from your web page.</span></span> <span data-ttu-id="e59d5-157">让我们修改示例代码，以使用 Excel 图标指示它启动 Excel 应用程序。</span><span class="sxs-lookup"><span data-stu-id="e59d5-157">Let’s modify the sample code to use the Excel icon to indicate that it launches the Excel application.</span></span>

1. <span data-ttu-id="e59d5-158">打开 Visual Studio 中的示例。</span><span class="sxs-lookup"><span data-stu-id="e59d5-158">Open the sample in Visual Studio.</span></span>
1. <span data-ttu-id="e59d5-159">打开 " **主页 .aspx** " 页面。</span><span class="sxs-lookup"><span data-stu-id="e59d5-159">Open the **Home.aspx** page.</span></span>
1. <span data-ttu-id="e59d5-160">查找以下代码，它是表单上的 "下载" 按钮：</span><span class="sxs-lookup"><span data-stu-id="e59d5-160">Find following code that is the download button on the form:</span></span>
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. <span data-ttu-id="e59d5-161">将按钮代码替换为以下图像标记。</span><span class="sxs-lookup"><span data-stu-id="e59d5-161">Replace the button code with the following image tag.</span></span>
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. <span data-ttu-id="e59d5-162">按 **F5** (或 **调试 > 启动调试**) 。</span><span class="sxs-lookup"><span data-stu-id="e59d5-162">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="e59d5-163">加载主页时，您会看到显示的图标。</span><span class="sxs-lookup"><span data-stu-id="e59d5-163">You'll see the icon appear when the home page loads.</span></span>

<span data-ttu-id="e59d5-164">有关详细信息，请参阅熟知的 UI 开发人员门户上的 [Office 品牌图标](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) 。</span><span class="sxs-lookup"><span data-stu-id="e59d5-164">For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.</span></span>  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a><span data-ttu-id="e59d5-165">将 Excel 文档上载到 Microsoft OneDrive</span><span class="sxs-lookup"><span data-stu-id="e59d5-165">Upload the Excel document to Microsoft OneDrive</span></span>

<span data-ttu-id="e59d5-166">如果客户使用 OneDrive，建议将新文档上载到 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="e59d5-166">We recommend uploading new documents to OneDrive if your customer uses OneDrive.</span></span> <span data-ttu-id="e59d5-167">这样，他们就可以更轻松地查找和使用文档。</span><span class="sxs-lookup"><span data-stu-id="e59d5-167">This makes it easier for them to find and work with the documents.</span></span> <span data-ttu-id="e59d5-168">我们来创建一个新的代码示例，并了解如何使用 Microsoft Graph SDK 将新的 Excel 文档上载到 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="e59d5-168">Let’s create a new code sample and see how you can use the Microsoft Graph SDK to upload a new Excel document to OneDrive.</span></span>

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a><span data-ttu-id="e59d5-169">使用快速启动构建新的 Microsoft Graph web 应用程序</span><span class="sxs-lookup"><span data-stu-id="e59d5-169">Use a quick-start to build a new Microsoft Graph web application</span></span>

1. <span data-ttu-id="e59d5-170">转到 [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 并按照步骤操作，以创建和打开与 Office 365 服务交互的快速入门代码示例。</span><span class="sxs-lookup"><span data-stu-id="e59d5-170">Go to [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) and follow the steps to create and open a quick start code sample that interacts with Office 365 services.</span></span>
1. <span data-ttu-id="e59d5-171">在 **步骤1：选择 "语言" 或 "平台**" 中，选择 " **ASP.NET MVC**"。</span><span class="sxs-lookup"><span data-stu-id="e59d5-171">In **step 1: Pick you language or platform**, choose **ASP.NET MVC**.</span></span> <span data-ttu-id="e59d5-172">虽然此过程中的步骤使用 ASP.NET MVC 选项，但这些步骤遵循适用于任何语言或平台的模式。</span><span class="sxs-lookup"><span data-stu-id="e59d5-172">Although the steps in this procedure use the ASP.NET MVC option, the steps follow a pattern that apply to any language or platform.</span></span>
1. <span data-ttu-id="e59d5-173">在 " **步骤2：获取应用程序 id 和密码**" 中，选择 " **获取应用 id 和密码**"。</span><span class="sxs-lookup"><span data-stu-id="e59d5-173">In **step 2: Get an app ID and secret**, choose **Get an app ID and secret**.</span></span>
1. <span data-ttu-id="e59d5-174">登录到 Microsoft 365 帐户。</span><span class="sxs-lookup"><span data-stu-id="e59d5-174">Sign in to your Microsoft 365 account.</span></span>  
1. <span data-ttu-id="e59d5-175">在 " **请保存您的应用程序机密** 网页" 中，将应用程序密码保存到文件位置，稍后可对其进行检索和使用。</span><span class="sxs-lookup"><span data-stu-id="e59d5-175">On the **Please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.</span></span>
1. <span data-ttu-id="e59d5-176">选择 **"已收到"，让我回到 "快速入门"**。</span><span class="sxs-lookup"><span data-stu-id="e59d5-176">Choose **Got it, take me back to the quick start**.</span></span>
1. <span data-ttu-id="e59d5-177">在 **第2步：注册成功！**</span><span class="sxs-lookup"><span data-stu-id="e59d5-177">In **step 2: Registration Successful!**</span></span> <span data-ttu-id="e59d5-178">输入生成的应用密码。</span><span class="sxs-lookup"><span data-stu-id="e59d5-178">Enter the generated app secret.</span></span>
1. <span data-ttu-id="e59d5-179">在 " **步骤3：开始编码**" 中，选择 **"下载基于 SDK 的代码" 示例**。</span><span class="sxs-lookup"><span data-stu-id="e59d5-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span></span>
1. <span data-ttu-id="e59d5-180">将下载 zip 文件夹解压缩到本地文件夹中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-180">Extract the download zip folder into a local folder.</span></span>  
1. <span data-ttu-id="e59d5-181">在 Visual Studio 2019 中打开 graph-tutorial 文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-181">Open the graph-tutorial.sln file in Visual Studio 2019.</span></span>
1. <span data-ttu-id="e59d5-182">生成并运行解决方案，并确认它是否正常工作。</span><span class="sxs-lookup"><span data-stu-id="e59d5-182">Build and run the solution and confirm it is working correctly.</span></span> <span data-ttu-id="e59d5-183">您应该能够使用 "日历" 网页查看您的 Microsoft 365 日历。</span><span class="sxs-lookup"><span data-stu-id="e59d5-183">You should be able to use the calendar web page to view your Microsoft 365 calendar.</span></span>

### <a name="upload-a-file-to-onedrive"></a><span data-ttu-id="e59d5-184">将文件上传到 OneDrive</span><span class="sxs-lookup"><span data-stu-id="e59d5-184">Upload a file to OneDrive</span></span>

1. <span data-ttu-id="e59d5-185">打开 Visual Studio 2019 中的 **graph-tutorial** 解决方案，然后打开 **PrivateSettings.config** 文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-185">Open the **graph-tutorial.sln** solution in Visual Studio 2019, and open the **PrivateSettings.config** file.</span></span>
1. <span data-ttu-id="e59d5-186">将新的作用域**文件**添加   到**ida： AppScopes**键，使其类似于以下代码：</span><span class="sxs-lookup"><span data-stu-id="e59d5-186">Add a new scope **Files.ReadWrite** to the **ida:AppScopes** key so that it looks like the following code:</span></span>
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. <span data-ttu-id="e59d5-187">打开 " **索引 cshtml** " 文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-187">Open the **Index.cshtml** file.</span></span>
1. <span data-ttu-id="e59d5-188">插入以下的 Html.actionlink 代码以创建按钮，以将文件上传到 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="e59d5-188">Insert the following ActionLink code to create a button to upload a file to OneDrive.</span></span>
    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```
1. <span data-ttu-id="e59d5-189">打开 **HomeController.cs** 文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-189">Open the **HomeController.cs** file.</span></span>
1. <span data-ttu-id="e59d5-190">插入以下代码以处理来自操作链接的请求。</span><span class="sxs-lookup"><span data-stu-id="e59d5-190">Insert the following code to handle the request from the action link.</span></span>
    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```
1. <span data-ttu-id="e59d5-191">打开 **GraphHelper.cs** 文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-191">Open the **GraphHelper.cs** file.</span></span>
1. <span data-ttu-id="e59d5-192">插入以下代码以调用 Microsoft Graph API，以在 OneDrive 上创建新文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-192">Insert the following code to call the Microsoft Graph API to create a new file on OneDrive.</span></span>
    ```csharp
    public static async Task UploadFile(string fileName, System.IO.MemoryStream stream)
        {
           var graphClient = GetAuthenticatedClient();
            await graphClient.Me
                .Drive
                .Root
                .ItemWithPath(fileName)
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);
            return;
        }
    ```
1. <span data-ttu-id="e59d5-193">按 **F5** (或 **调试 > 启动调试**) 。</span><span class="sxs-lookup"><span data-stu-id="e59d5-193">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="e59d5-194">Web 应用程序将启动。</span><span class="sxs-lookup"><span data-stu-id="e59d5-194">The web application will start.</span></span>
1. <span data-ttu-id="e59d5-195">选择 **"单击此处登录**"，然后登录。</span><span class="sxs-lookup"><span data-stu-id="e59d5-195">Choose **Click here to sign in**, and sign in.</span></span>
1. <span data-ttu-id="e59d5-196">选择 " **单击此处可在 OneDrive 上创建新文件"**。</span><span class="sxs-lookup"><span data-stu-id="e59d5-196">Choose **Click here to create a new file on OneDrive**.</span></span>
1. <span data-ttu-id="e59d5-197">打开一个新的浏览器选项卡并登录到您的 OneDrive 帐户。</span><span class="sxs-lookup"><span data-stu-id="e59d5-197">Open a new browser tab and sign in to your OneDrive account.</span></span> <span data-ttu-id="e59d5-198">您将在根文件夹中看到 test.txt 文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-198">You'll see the test.txt file in the root folder.</span></span>

<span data-ttu-id="e59d5-199">现在，您已经了解如何将文件上传到 OneDrive，您可以重复使用此代码来上传您创建的任何 Excel 文档。</span><span class="sxs-lookup"><span data-stu-id="e59d5-199">Now that you've learned how to upload a file to OneDrive, you can reuse this code to upload any Excel document that you create.</span></span>

## <a name="additional-considerations-for-your-solution"></a><span data-ttu-id="e59d5-200">解决方案的其他注意事项</span><span class="sxs-lookup"><span data-stu-id="e59d5-200">Additional considerations for your solution</span></span>

<span data-ttu-id="e59d5-201">每个人的解决方案在技术和方法方面各不相同。</span><span class="sxs-lookup"><span data-stu-id="e59d5-201">Everyone’s solution is different in terms of technologies and approaches.</span></span> <span data-ttu-id="e59d5-202">以下注意事项将帮助您规划如何修改解决方案以打开文档并嵌入 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="e59d5-202">The following considerations will help you plan how to modify your solution to open documents and embed your Office Add-in.</span></span>

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a><span data-ttu-id="e59d5-203">从网页创建新 Excel 电子表格</span><span class="sxs-lookup"><span data-stu-id="e59d5-203">Create a new Excel spreadsheet from the web page</span></span>

<span data-ttu-id="e59d5-204">此示例修改现有的 Excel 文档。</span><span class="sxs-lookup"><span data-stu-id="e59d5-204">The sample modifies an existing Excel document.</span></span> <span data-ttu-id="e59d5-205">一个更常见的方案是，从网页创建一个新的 Excel 电子表格。</span><span class="sxs-lookup"><span data-stu-id="e59d5-205">A more common scenario is that you’ll create a new Excel spreadsheet from your web page.</span></span> <span data-ttu-id="e59d5-206">您可以通过提供文件名来查找有关如何在 **创建电子表格文档** 中创建新电子表格的其他详细信息。</span><span class="sxs-lookup"><span data-stu-id="e59d5-206">You can find additional details on how to create a new spreadsheet in **Create a spreadsheet document** by providing a file name.</span></span> <span data-ttu-id="e59d5-207">本文介绍如何在本地创建文件，但您也可以使用 SpreadsheetDocument 方法上的重载在 stream 中创建文件。</span><span class="sxs-lookup"><span data-stu-id="e59d5-207">This article shows how to create the file locally, but you can also create the file in a stream by using an overload on the SpreadsheetDocument.Create method.</span></span>

### <a name="read-custom-properties-when-your-add-in-starts"></a><span data-ttu-id="e59d5-208">在外接程序启动时读取自定义属性</span><span class="sxs-lookup"><span data-stu-id="e59d5-208">Read custom properties when your add-in starts</span></span>

<span data-ttu-id="e59d5-209">该代码示例使用 OOXML SDK 将一个代码段 ID 存储在新的 Excel 文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-209">The code sample stores a snippet ID in the new Excel document using the OOXML SDK.</span></span> <span data-ttu-id="e59d5-210">脚本实验室从 Excel 文档读取代码段 ID，然后在它打开时显示该代码段。</span><span class="sxs-lookup"><span data-stu-id="e59d5-210">Script Lab reads the snippet ID from the Excel document and then displays that snippet code when it opens.</span></span> <span data-ttu-id="e59d5-211">您可能需要将自定义属性发送到您自己的外接程序 (例如查询字符串或临时身份验证令牌。 ) 请参阅 **保留外接程序状态和设置** ，了解有关如何在加载项启动时读取自定义属性的完整详细信息。</span><span class="sxs-lookup"><span data-stu-id="e59d5-211">You may need to send custom properties to your own add-in (such as a query string, or temporary authentication token.) See **Persisting add-in state and settings** for complete details on how to read custom properties when your add-in starts.</span></span>

### <a name="initialize-the-excel-document-with-data"></a><span data-ttu-id="e59d5-212">使用数据初始化 Excel 文档</span><span class="sxs-lookup"><span data-stu-id="e59d5-212">Initialize the Excel document with data</span></span>

<span data-ttu-id="e59d5-213">通常，当客户从您的网站打开 Excel 文档时，他们希望文档包含网站中的一些数据。</span><span class="sxs-lookup"><span data-stu-id="e59d5-213">Typically, when the customer opens up an Excel document from your web site, they expect the document to contain some data from the web site.</span></span> <span data-ttu-id="e59d5-214">有几种方法可将数据写入文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-214">There are a couple of ways to write data into the document.</span></span>

- <span data-ttu-id="e59d5-215">**使用 OOXML SDK 写入数据**。</span><span class="sxs-lookup"><span data-stu-id="e59d5-215">**Use the OOXML SDK to write the data**.</span></span> <span data-ttu-id="e59d5-216">您可以使用 SDK 直接将任何数据写入文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-216">You can use the SDK to directly write any data into the document.</span></span> <span data-ttu-id="e59d5-217">如果您希望数据在文档打开时即时可用，则此方法很有用。</span><span class="sxs-lookup"><span data-stu-id="e59d5-217">This approach is useful if you want the data to be available the instant the document is opened.</span></span>
- <span data-ttu-id="e59d5-218">将**自定义查询属性传递到 Office 外接程序**。</span><span class="sxs-lookup"><span data-stu-id="e59d5-218">**Pass a custom query property to your Office add-in**.</span></span> <span data-ttu-id="e59d5-219">在生成文档时，您嵌入了 Office 加载项的自定义属性，该属性包含检索所有必需数据的查询字符串。</span><span class="sxs-lookup"><span data-stu-id="e59d5-219">When you generate the document, you embed a custom property for the Office add-in that contains a query string that retrieves all the required data.</span></span> <span data-ttu-id="e59d5-220">当您的外接程序打开时，它将检索查询，运行查询，并使用 Office JS API 将查询结果插入到文档中。</span><span class="sxs-lookup"><span data-stu-id="e59d5-220">When your add-in opens, it retrieves the query, runs the query, and uses the Office JS API to insert the result of the query into the document.</span></span>

### <a name="working-with-the-ooxml-sdk"></a><span data-ttu-id="e59d5-221">使用 OOXML SDK</span><span class="sxs-lookup"><span data-stu-id="e59d5-221">Working with the OOXML SDK</span></span>

<span data-ttu-id="e59d5-222">OOXML SDK 基于 .NET。</span><span class="sxs-lookup"><span data-stu-id="e59d5-222">The OOXML SDK is based on .NET.</span></span> <span data-ttu-id="e59d5-223">如果您的 web 应用程序不是 .NET，则需要查找另一种使用 OOXML 的方法。</span><span class="sxs-lookup"><span data-stu-id="e59d5-223">If your web application does not .NET, you’ll need to look for an alternative way to work with OOXML.</span></span>

<span data-ttu-id="e59d5-224">在适用于 [javascript 的 OPEN XML SDK](https://archive.codeplex.com/?p=openxmlsdkjs)中，有一个适用于 OOXML Sdk 的 JavaScript 版本。</span><span class="sxs-lookup"><span data-stu-id="e59d5-224">There is a JavaScript version of the OOXML SDK available at [Open XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span></span>

<span data-ttu-id="e59d5-225">您可以将 OOXML 代码放在 Azure 函数中，以将 .NET 代码与 web 应用程序的其余部分分开。</span><span class="sxs-lookup"><span data-stu-id="e59d5-225">You can place the OOXML code in an Azure function to separate the .NET code from the rest of your web application.</span></span> <span data-ttu-id="e59d5-226">然后，调用 Azure 函数 (从 Web 应用程序生成 Excel 文档) 。</span><span class="sxs-lookup"><span data-stu-id="e59d5-226">Then call the Azure function (to generate the Excel document) from your Web application.</span></span> <span data-ttu-id="e59d5-227">有关 Azure 函数的详细信息，请参阅 [Azure 函数简介](https://docs.microsoft.com/azure/azure-functions/functions-overview)。</span><span class="sxs-lookup"><span data-stu-id="e59d5-227">For more information on Azure functions, see [An introduction to Azure Functions](https://docs.microsoft.com/azure/azure-functions/functions-overview).</span></span>

### <a name="use-single-sign-on"></a><span data-ttu-id="e59d5-228">使用单一登录</span><span class="sxs-lookup"><span data-stu-id="e59d5-228">Use single sign-on</span></span>

<span data-ttu-id="e59d5-229">为了简化身份验证，我们建议你的外接程序实现单一登录。</span><span class="sxs-lookup"><span data-stu-id="e59d5-229">To simplify authentication, we recommend your add-in implements single sign-on.</span></span> <span data-ttu-id="e59d5-230">有关详细信息，请参阅 [为 Office 加载项启用单一登录](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="e59d5-230">For more information, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="e59d5-231">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e59d5-231">See also</span></span>

- [<span data-ttu-id="e59d5-232">欢迎使用 Open XML SDK 2.5 for Office</span><span class="sxs-lookup"><span data-stu-id="e59d5-232">Welcome to the Open XML SDK 2.5 for Office</span></span>](/office/open-xml/open-xml-sdk)
- [<span data-ttu-id="e59d5-233">随文档自动打开任务窗格</span><span class="sxs-lookup"><span data-stu-id="e59d5-233">Automatically open a task pane with a document</span></span>](../develop/automatically-open-a-task-pane-with-a-document.md)
- [<span data-ttu-id="e59d5-234">保留加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="e59d5-234">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="e59d5-235">通过提供文件名创建电子表格文档</span><span class="sxs-lookup"><span data-stu-id="e59d5-235">Create a spreadsheet document by providing a file name</span></span>](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)
