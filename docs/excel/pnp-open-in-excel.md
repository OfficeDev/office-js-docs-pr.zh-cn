---
title: 从Excel打开加载项并嵌入Office加载项
description: 从Excel打开"加载项"，并嵌入Office加载项。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 18f40b0030f4132a413a879e8b3419af49984b45
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349376"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a><span data-ttu-id="67a38-103">从Excel打开加载项并嵌入Office加载项</span><span class="sxs-lookup"><span data-stu-id="67a38-103">Open Excel from your web page and embed your Office Add-in</span></span>

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="网页上Excel按钮的图像，该按钮可打开一个新的 Excel 文档，并嵌入并自动打开外接程序。":::

<span data-ttu-id="67a38-105">扩展 SaaS Web 应用程序，以便客户可以直接在网页中打开其数据Microsoft Excel。</span><span class="sxs-lookup"><span data-stu-id="67a38-105">Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel.</span></span> <span data-ttu-id="67a38-106">一种常见方案是客户将处理 Web 应用程序中的数据。</span><span class="sxs-lookup"><span data-stu-id="67a38-106">A common scenario is that customers will be working with data in your web application.</span></span> <span data-ttu-id="67a38-107">然后，他们希望将数据复制到一个Excel文档中。</span><span class="sxs-lookup"><span data-stu-id="67a38-107">Then they’ll want to copy the data into an Excel document.</span></span> <span data-ttu-id="67a38-108">例如，他们可能需要使用数据进行其他Excel。</span><span class="sxs-lookup"><span data-stu-id="67a38-108">For example, they may want to perform additional analysis using Excel.</span></span> <span data-ttu-id="67a38-109">通常，客户需要将数据导出到文件（如 .csv 文件）中，然后将该数据导入Excel。</span><span class="sxs-lookup"><span data-stu-id="67a38-109">Typically, the customer is required to export the data to a file, such as a .csv file, and then import that data into Excel.</span></span> <span data-ttu-id="67a38-110">他们还必须手动将Office加载项添加到文档中。</span><span class="sxs-lookup"><span data-stu-id="67a38-110">They also have to manually add your Office Add-in to the document.</span></span>

<span data-ttu-id="67a38-111">将步骤数减少为生成文档并打开文档的网页上的单个Excel单击。</span><span class="sxs-lookup"><span data-stu-id="67a38-111">Reduce the number of steps to a single button click on your web page that generates and opens the Excel document.</span></span> <span data-ttu-id="67a38-112">您还可以在文档中Office外接程序，在文档打开时显示它。</span><span class="sxs-lookup"><span data-stu-id="67a38-112">You can also embed your Office Add-in inside the document and display it when the document opens.</span></span> <span data-ttu-id="67a38-113">这将确保客户仍可访问你的应用程序功能。</span><span class="sxs-lookup"><span data-stu-id="67a38-113">This ensures the customer still has access to your application features.</span></span> <span data-ttu-id="67a38-114">当文档打开时，客户选择的数据以及你的Office外接程序已可供他们继续工作。</span><span class="sxs-lookup"><span data-stu-id="67a38-114">When the document opens, the data the customer selected, and your Office Add-in is already available for them to continue working.</span></span>

<span data-ttu-id="67a38-115">本文介绍了在你自己的 SaaS Web 应用程序中实现此方案的代码和技术。</span><span class="sxs-lookup"><span data-stu-id="67a38-115">This article shows you code and techniques for implementing this scenario in your own SaaS web application.</span></span>

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a><span data-ttu-id="67a38-116">新建Excel文档并嵌入Office加载项</span><span class="sxs-lookup"><span data-stu-id="67a38-116">Create a new Excel document and embed an Office Add-in</span></span>

<span data-ttu-id="67a38-117">首先，让我们了解如何从网页Excel文档，以及如何在文档中嵌入加载项。</span><span class="sxs-lookup"><span data-stu-id="67a38-117">First, let’s learn how to create an Excel document from a web page, and embed an add-in into the document.</span></span> <span data-ttu-id="67a38-118">the [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the Script Lab [add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span><span class="sxs-lookup"><span data-stu-id="67a38-118">The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span></span> <span data-ttu-id="67a38-119">尽管该示例适用于Office文档，但我们将重点介绍本文Excel电子表格。</span><span class="sxs-lookup"><span data-stu-id="67a38-119">Although the sample works with any Office document, we’ll just focus on Excel spreadsheets in this article.</span></span> <span data-ttu-id="67a38-120">使用以下步骤生成并运行示例。</span><span class="sxs-lookup"><span data-stu-id="67a38-120">Use the following steps to build and run the sample.</span></span>

1. <span data-ttu-id="67a38-121">将示例代码从  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip 中提取到您计算机的文件夹中。</span><span class="sxs-lookup"><span data-stu-id="67a38-121">Extract the sample code from  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip into a folder on your computer.</span></span>
2. <span data-ttu-id="67a38-122">若要生成并运行示例，请按照自述文件" **使用项目"** 部分的步骤操作。</span><span class="sxs-lookup"><span data-stu-id="67a38-122">To build and run the sample, follow the steps in the **To use the project** section of the readme.</span></span>
3. <span data-ttu-id="67a38-123">运行示例时，将显示类似于以下屏幕截图的网页。</span><span class="sxs-lookup"><span data-stu-id="67a38-123">When you run the sample it will display a web page similar to the following screenshot.</span></span> <span data-ttu-id="67a38-124">使用网页创建一个新的Excel文档，其中包含Script Lab打开时的内容。</span><span class="sxs-lookup"><span data-stu-id="67a38-124">Use the web page to create a new Excel document that contains Script Lab when it opens.</span></span>
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="嵌入脚本实验室示例显示的网页的屏幕截图，用于选择Excel文件并将脚本实验室外接程序嵌入其中。":::

### <a name="how-the-sample-works"></a><span data-ttu-id="67a38-126">示例的工作原理</span><span class="sxs-lookup"><span data-stu-id="67a38-126">How the sample works</span></span>

<span data-ttu-id="67a38-127">示例代码使用 OOXML SDK 将Script Lab嵌入到您Excel文档。</span><span class="sxs-lookup"><span data-stu-id="67a38-127">The sample code uses the OOXML SDK to embed the Script Lab add-in to the Excel document that you choose.</span></span> <span data-ttu-id="67a38-128">以下信息来自自述 [**文件的关于**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md)代码部分。</span><span class="sxs-lookup"><span data-stu-id="67a38-128">The following information is taken from the [**About the code** section](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) in the readme file.</span></span>

<span data-ttu-id="67a38-129">文件 **Home.aspx.cs：**</span><span class="sxs-lookup"><span data-stu-id="67a38-129">The file **Home.aspx.cs**:</span></span>

- <span data-ttu-id="67a38-130">提供按钮事件处理程序和基本 UI 操作。</span><span class="sxs-lookup"><span data-stu-id="67a38-130">Provides the button event handlers and basic UI manipulation.</span></span>
- <span data-ttu-id="67a38-131">使用 ASP.NET 技术上载和下载文件。</span><span class="sxs-lookup"><span data-stu-id="67a38-131">Uses standard ASP.NET techniques to upload and download the file.</span></span>
- <span data-ttu-id="67a38-132">使用 xlsx、docx 或 pptx (上传的文件的文件扩展名) 确定文件类型。</span><span class="sxs-lookup"><span data-stu-id="67a38-132">Uses the file name extension of the uploaded file (xlsx, docx, or pptx) to determine the type of file.</span></span> <span data-ttu-id="67a38-133">需要从一开始就完成此操作，因为 Open XML SDK 通常对于每种类型的文件都有不同的 API。</span><span class="sxs-lookup"><span data-stu-id="67a38-133">This needs to be done at the outset because the Open XML SDK generally has distinct APIs for each type of file.</span></span>
- <span data-ttu-id="67a38-134">调用 **OOXMLHelper** 以验证文件，并调用 **AddInEmbedder** 以在Script Lab嵌入文件并设置为自动打开。</span><span class="sxs-lookup"><span data-stu-id="67a38-134">Calls into the **OOXMLHelper** to validate the file and calls into the **AddInEmbedder** to embed Script Lab in the file and set to automatically open.</span></span>

<span data-ttu-id="67a38-135">文件 **AddInEmbedder.cs**：</span><span class="sxs-lookup"><span data-stu-id="67a38-135">The file **AddInEmbedder.cs**:</span></span>

- <span data-ttu-id="67a38-136">提供主要业务逻辑，此示例中是嵌入 Script Lab。</span><span class="sxs-lookup"><span data-stu-id="67a38-136">Provides the main business logic, which in this sample is a method that embeds Script Lab.</span></span>
- <span data-ttu-id="67a38-137">根据文件类型调用 OOXML 帮助程序。</span><span class="sxs-lookup"><span data-stu-id="67a38-137">Makes calls into the OOXML helper based on the type of the file.</span></span>

<span data-ttu-id="67a38-138">文件 **OOXMLHelper.cs：**</span><span class="sxs-lookup"><span data-stu-id="67a38-138">The file **OOXMLHelper.cs**:</span></span>

- <span data-ttu-id="67a38-139">提供所有详细的 OOXML 操作。</span><span class="sxs-lookup"><span data-stu-id="67a38-139">Provides all the detailed OOXML manipulation.</span></span>
- <span data-ttu-id="67a38-140">使用标准技术来验证Office文件，只需对该文件 **调用 Document.Open** 方法。</span><span class="sxs-lookup"><span data-stu-id="67a38-140">Uses a standard technique for validating the Office file, which is simply to call the **Document.Open** method on it.</span></span> <span data-ttu-id="67a38-141">如果文件无效，该方法将引发异常。</span><span class="sxs-lookup"><span data-stu-id="67a38-141">If the file is invalid, the method throws an exception.</span></span>
- <span data-ttu-id="67a38-142">包含主要由 Open XML 2.5 SDK Productivity Tools 生成的代码，这些代码位于 [Open XML 2.5 SDK 的链接中](/office/open-xml/open-xml-sdk)。</span><span class="sxs-lookup"><span data-stu-id="67a38-142">Contains mainly code that was generated by the Open XML 2.5 SDK Productivity Tools which are available at the link for the [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk).</span></span>

<span data-ttu-id="67a38-143">**OOXMLHelper.cs** 文件中 **GenerateWebExtensionPart1Content** 方法设置对 Microsoft AppSource 中 Script Lab ID 的引用：</span><span class="sxs-lookup"><span data-stu-id="67a38-143">The **GenerateWebExtensionPart1Content** method in the **OOXMLHelper.cs** file sets the reference to the ID of Script Lab in Microsoft AppSource:</span></span>

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- <span data-ttu-id="67a38-144">**StoreType** 值为"OMEX"，它是 Microsoft AppSource 的别名。</span><span class="sxs-lookup"><span data-stu-id="67a38-144">The **StoreType** value is "OMEX", an alias for Microsoft AppSource.</span></span>
- <span data-ttu-id="67a38-145">Store 值为"en-US"，可以在 Microsoft AppSource 区域性部分找到Script Lab。</span><span class="sxs-lookup"><span data-stu-id="67a38-145">The **Store** value is "en-US" found in the Microsoft AppSource culture section for Script Lab.</span></span>
- <span data-ttu-id="67a38-146">**Id** 值是 Microsoft AppSource 资产 ID Script Lab。</span><span class="sxs-lookup"><span data-stu-id="67a38-146">The **Id** value is the Microsoft AppSource asset ID for Script Lab.</span></span>

<span data-ttu-id="67a38-147">如果要从文件共享目录设置外接程序以自动打开，你将使用不同的值：</span><span class="sxs-lookup"><span data-stu-id="67a38-147">If you are setting up an add-in from a file share catalog for auto-open, you will use different values:</span></span>

<span data-ttu-id="67a38-148">**StoreType** 值为"FileSystem"。</span><span class="sxs-lookup"><span data-stu-id="67a38-148">The **StoreType** value is "FileSystem".</span></span>

- <span data-ttu-id="67a38-149">**Store** 值是网络共享 URL;例如 \\ \\ ，"MyComputer \\ MySharedFolder"。</span><span class="sxs-lookup"><span data-stu-id="67a38-149">The **Store** value is the URL of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span> <span data-ttu-id="67a38-150">这应该是在共享信任中心显示为共享受信任目录地址Office URL。</span><span class="sxs-lookup"><span data-stu-id="67a38-150">This should be the exact URL that appears as the share's Trusted Catalog Address in the Office Trust Center.</span></span>
- <span data-ttu-id="67a38-151">**Id** 值是外接程序清单中的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="67a38-151">The **Id** value is the app ID in the add-ins manifest.</span></span>
> [!NOTE]
> <span data-ttu-id="67a38-152">有关这些属性的可选值的详细信息，请参阅 [使用文档自动打开任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)。</span><span class="sxs-lookup"><span data-stu-id="67a38-152">For more information about alternative values for these attributes, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).</span></span>

## <a name="use-the-fluent-ui"></a><span data-ttu-id="67a38-153">使用 Fluent UI</span><span class="sxs-lookup"><span data-stu-id="67a38-153">Use the Fluent UI</span></span>

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="FluentWord、Excel 和 PowerPoint 的 UI 图标。":::

<span data-ttu-id="67a38-155">最佳做法是使用 Fluent UI 来帮助用户在 Microsoft 产品之间过渡。</span><span class="sxs-lookup"><span data-stu-id="67a38-155">A best practice is to use the Fluent UI to help your users transition between Microsoft products.</span></span> <span data-ttu-id="67a38-156">应始终使用Office图标来指示Office从网页启动哪个应用程序。</span><span class="sxs-lookup"><span data-stu-id="67a38-156">You should always use an Office icon to indicate which Office application will be launched from your web page.</span></span> <span data-ttu-id="67a38-157">让我们修改示例代码，以使用 Excel 图标指示它启动 Excel 应用程序。</span><span class="sxs-lookup"><span data-stu-id="67a38-157">Let’s modify the sample code to use the Excel icon to indicate that it launches the Excel application.</span></span>

1. <span data-ttu-id="67a38-158">在"管理"中Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="67a38-158">Open the sample in Visual Studio.</span></span>
1. <span data-ttu-id="67a38-159">打开 **Home.aspx** 页。</span><span class="sxs-lookup"><span data-stu-id="67a38-159">Open the **Home.aspx** page.</span></span>
1. <span data-ttu-id="67a38-160">在表单上查找以下作为下载按钮的代码。</span><span class="sxs-lookup"><span data-stu-id="67a38-160">Find following code that is the download button on the form.</span></span>

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. <span data-ttu-id="67a38-161">将按钮代码替换为以下图像标记。</span><span class="sxs-lookup"><span data-stu-id="67a38-161">Replace the button code with the following image tag.</span></span>

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. <span data-ttu-id="67a38-162">按 **F5** (**或调试>开始调试**) 。</span><span class="sxs-lookup"><span data-stu-id="67a38-162">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="67a38-163">加载主页时，你将看到图标出现。</span><span class="sxs-lookup"><span data-stu-id="67a38-163">You'll see the icon appear when the home page loads.</span></span>

<span data-ttu-id="67a38-164">有关详细信息，请参阅[Office](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) UI 开发人员门户Fluent品牌图标。</span><span class="sxs-lookup"><span data-stu-id="67a38-164">For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.</span></span>  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a><span data-ttu-id="67a38-165">Upload Excel文档Microsoft OneDrive</span><span class="sxs-lookup"><span data-stu-id="67a38-165">Upload the Excel document to Microsoft OneDrive</span></span>

<span data-ttu-id="67a38-166">如果你的客户使用 OneDrive，我们建议将新文档OneDrive。</span><span class="sxs-lookup"><span data-stu-id="67a38-166">We recommend uploading new documents to OneDrive if your customer uses OneDrive.</span></span> <span data-ttu-id="67a38-167">这使用户更易于查找并处理文档。</span><span class="sxs-lookup"><span data-stu-id="67a38-167">This makes it easier for them to find and work with the documents.</span></span> <span data-ttu-id="67a38-168">让我们创建新的代码示例，了解如何使用 Microsoft Graph SDK 将新的 Excel 文档上载到OneDrive。</span><span class="sxs-lookup"><span data-stu-id="67a38-168">Let’s create a new code sample and see how you can use the Microsoft Graph SDK to upload a new Excel document to OneDrive.</span></span>

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a><span data-ttu-id="67a38-169">使用快速入门生成新的 Microsoft Graph Web 应用程序</span><span class="sxs-lookup"><span data-stu-id="67a38-169">Use a quick-start to build a new Microsoft Graph web application</span></span>

1. <span data-ttu-id="67a38-170">转到 [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 并按照步骤创建并打开与服务交互的快速启动Office示例。</span><span class="sxs-lookup"><span data-stu-id="67a38-170">Go to [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) and follow the steps to create and open a quick start code sample that interacts with Office services.</span></span>
1. <span data-ttu-id="67a38-171">在 **步骤 1：选择语言或平台中**，选择 **"ASP.NET MVC"。**</span><span class="sxs-lookup"><span data-stu-id="67a38-171">In **step 1: Pick you language or platform**, choose **ASP.NET MVC**.</span></span> <span data-ttu-id="67a38-172">虽然此过程中的步骤使用 ASP.NET MVC 选项，但步骤遵循适用于任何语言或平台的模式。</span><span class="sxs-lookup"><span data-stu-id="67a38-172">Although the steps in this procedure use the ASP.NET MVC option, the steps follow a pattern that apply to any language or platform.</span></span>
1. <span data-ttu-id="67a38-173">在 **步骤 2：获取应用 ID 和密码中，** 选择 **"获取应用 ID 和密码"。**</span><span class="sxs-lookup"><span data-stu-id="67a38-173">In **step 2: Get an app ID and secret**, choose **Get an app ID and secret**.</span></span>
1. <span data-ttu-id="67a38-174">登录到你的 Microsoft 365 帐户。</span><span class="sxs-lookup"><span data-stu-id="67a38-174">Sign in to your Microsoft 365 account.</span></span>  
1. <span data-ttu-id="67a38-175">在 **"请保存应用密码** "网页上，将应用密码保存到稍后可以检索和使用的文件位置。</span><span class="sxs-lookup"><span data-stu-id="67a38-175">On the **Please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.</span></span>
1. <span data-ttu-id="67a38-176">选择 **"已接受"，将我返回到快速入门**。</span><span class="sxs-lookup"><span data-stu-id="67a38-176">Choose **Got it, take me back to the quick start**.</span></span>
1. <span data-ttu-id="67a38-177">在 **步骤 2：注册成功！**</span><span class="sxs-lookup"><span data-stu-id="67a38-177">In **step 2: Registration Successful!**</span></span> <span data-ttu-id="67a38-178">输入生成的应用密码。</span><span class="sxs-lookup"><span data-stu-id="67a38-178">Enter the generated app secret.</span></span>
1. <span data-ttu-id="67a38-179">在 **"步骤 3： 开始编码"中**，**选择"下载基于 SDK 的代码示例"。**</span><span class="sxs-lookup"><span data-stu-id="67a38-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span></span>
1. <span data-ttu-id="67a38-180">将下载 zip 文件夹解压缩到本地文件夹。</span><span class="sxs-lookup"><span data-stu-id="67a38-180">Extract the download zip folder into a local folder.</span></span>  
1. <span data-ttu-id="67a38-181">在 2019 年 10 月Visual Studio graph-tutorial.sln 文件。</span><span class="sxs-lookup"><span data-stu-id="67a38-181">Open the graph-tutorial.sln file in Visual Studio 2019.</span></span>
1. <span data-ttu-id="67a38-182">生成并运行解决方案并确认它正常工作。</span><span class="sxs-lookup"><span data-stu-id="67a38-182">Build and run the solution and confirm it is working correctly.</span></span> <span data-ttu-id="67a38-183">您应该能够使用日历网页来查看您的日历Microsoft 365日历。</span><span class="sxs-lookup"><span data-stu-id="67a38-183">You should be able to use the calendar web page to view your Microsoft 365 calendar.</span></span>

### <a name="upload-a-file-to-onedrive"></a><span data-ttu-id="67a38-184">Upload文件OneDrive</span><span class="sxs-lookup"><span data-stu-id="67a38-184">Upload a file to OneDrive</span></span>

1. <span data-ttu-id="67a38-185">在 Visual Studio 2019 中打开 **graph-tutorial.sln** 解决方案，PrivateSettings.config **文件。**</span><span class="sxs-lookup"><span data-stu-id="67a38-185">Open the **graph-tutorial.sln** solution in Visual Studio 2019, and open the **PrivateSettings.config** file.</span></span>
1. <span data-ttu-id="67a38-186">将新的作用域 **Files.ReadWrite**   添加到 **ida：AppScopes** 项，以便它类似于以下代码。</span><span class="sxs-lookup"><span data-stu-id="67a38-186">Add a new scope **Files.ReadWrite** to the **ida:AppScopes** key so that it looks like the following code.</span></span>

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. <span data-ttu-id="67a38-187">打开 **Index.cshtml** 文件。</span><span class="sxs-lookup"><span data-stu-id="67a38-187">Open the **Index.cshtml** file.</span></span>
1. <span data-ttu-id="67a38-188">插入以下 ActionLink 代码以创建一个按钮以将文件上载到OneDrive。</span><span class="sxs-lookup"><span data-stu-id="67a38-188">Insert the following ActionLink code to create a button to upload a file to OneDrive.</span></span>

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. <span data-ttu-id="67a38-189">打开 **HomeController.cs** 文件。</span><span class="sxs-lookup"><span data-stu-id="67a38-189">Open the **HomeController.cs** file.</span></span>
1. <span data-ttu-id="67a38-190">插入以下代码以处理来自操作链接的请求。</span><span class="sxs-lookup"><span data-stu-id="67a38-190">Insert the following code to handle the request from the action link.</span></span>

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. <span data-ttu-id="67a38-191">打开 **GraphHelper.cs** 文件。</span><span class="sxs-lookup"><span data-stu-id="67a38-191">Open the **GraphHelper.cs** file.</span></span>
1. <span data-ttu-id="67a38-192">插入以下代码以调用 Microsoft Graph API，以在 OneDrive。</span><span class="sxs-lookup"><span data-stu-id="67a38-192">Insert the following code to call the Microsoft Graph API to create a new file on OneDrive.</span></span>

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

1. <span data-ttu-id="67a38-193">按 **F5** (**或调试>开始调试**) 。</span><span class="sxs-lookup"><span data-stu-id="67a38-193">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="67a38-194">Web 应用程序将启动。</span><span class="sxs-lookup"><span data-stu-id="67a38-194">The web application will start.</span></span>
1. <span data-ttu-id="67a38-195">选择 **"单击此处登录"，** 然后登录。</span><span class="sxs-lookup"><span data-stu-id="67a38-195">Choose **Click here to sign in**, and sign in.</span></span>
1. <span data-ttu-id="67a38-196">选择 **"单击此处"以在"新建OneDrive"。**</span><span class="sxs-lookup"><span data-stu-id="67a38-196">Choose **Click here to create a new file on OneDrive**.</span></span>
1. <span data-ttu-id="67a38-197">打开新的浏览器选项卡，然后登录你的OneDrive帐户。</span><span class="sxs-lookup"><span data-stu-id="67a38-197">Open a new browser tab and sign in to your OneDrive account.</span></span> <span data-ttu-id="67a38-198">你将看到根文件夹中test.txt文件。</span><span class="sxs-lookup"><span data-stu-id="67a38-198">You'll see the test.txt file in the root folder.</span></span>

<span data-ttu-id="67a38-199">现在，你已了解如何将文件上载到OneDrive，可以重复使用此代码上载Excel创建的任何文档。</span><span class="sxs-lookup"><span data-stu-id="67a38-199">Now that you've learned how to upload a file to OneDrive, you can reuse this code to upload any Excel document that you create.</span></span>

## <a name="additional-considerations-for-your-solution"></a><span data-ttu-id="67a38-200">解决方案的其他注意事项</span><span class="sxs-lookup"><span data-stu-id="67a38-200">Additional considerations for your solution</span></span>

<span data-ttu-id="67a38-201">每个人的解决方案在技术和方法方面是不同的。</span><span class="sxs-lookup"><span data-stu-id="67a38-201">Everyone’s solution is different in terms of technologies and approaches.</span></span> <span data-ttu-id="67a38-202">以下注意事项将帮助您规划如何修改解决方案以打开文档并嵌入Office外接程序。</span><span class="sxs-lookup"><span data-stu-id="67a38-202">The following considerations will help you plan how to modify your solution to open documents and embed your Office Add-in.</span></span>

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a><span data-ttu-id="67a38-203">从网页Excel新建一个电子表格</span><span class="sxs-lookup"><span data-stu-id="67a38-203">Create a new Excel spreadsheet from the web page</span></span>

<span data-ttu-id="67a38-204">本示例修改现有文档Excel文档。</span><span class="sxs-lookup"><span data-stu-id="67a38-204">The sample modifies an existing Excel document.</span></span> <span data-ttu-id="67a38-205">更常见的方案是，从网页Excel一个新的电子表格。</span><span class="sxs-lookup"><span data-stu-id="67a38-205">A more common scenario is that you’ll create a new Excel spreadsheet from your web page.</span></span> <span data-ttu-id="67a38-206">在"通过提供文件名创建电子表格文档"中，可以找到 **有关新建电子表格** 的其他详细信息。</span><span class="sxs-lookup"><span data-stu-id="67a38-206">You can find additional details on how to create a new spreadsheet in **Create a spreadsheet document** by providing a file name.</span></span> <span data-ttu-id="67a38-207">本文演示如何在本地创建文件，但您也可以使用 SpreadsheetDocument.Create 方法上的重载在流中创建文件。</span><span class="sxs-lookup"><span data-stu-id="67a38-207">This article shows how to create the file locally, but you can also create the file in a stream by using an overload on the SpreadsheetDocument.Create method.</span></span>

### <a name="read-custom-properties-when-your-add-in-starts"></a><span data-ttu-id="67a38-208">在加载项启动时读取自定义属性</span><span class="sxs-lookup"><span data-stu-id="67a38-208">Read custom properties when your add-in starts</span></span>

<span data-ttu-id="67a38-209">该代码示例使用 OOXML SDK 将代码段 ID Excel文档。</span><span class="sxs-lookup"><span data-stu-id="67a38-209">The code sample stores a snippet ID in the new Excel document using the OOXML SDK.</span></span> <span data-ttu-id="67a38-210">Script Lab从文档读取代码Excel ID，然后在代码段打开时显示该代码段。</span><span class="sxs-lookup"><span data-stu-id="67a38-210">Script Lab reads the snippet ID from the Excel document and then displays that snippet code when it opens.</span></span> <span data-ttu-id="67a38-211">您可能需要将自定义属性发送到您自己的外接程序 (例如查询字符串或临时身份验证令牌。) 请参阅持久化外接程序状态和设置，了解有关在外接程序启动时如何读取自定义属性的完整详细信息。</span><span class="sxs-lookup"><span data-stu-id="67a38-211">You may need to send custom properties to your own add-in (such as a query string, or temporary authentication token.) See **Persisting add-in state and settings** for complete details on how to read custom properties when your add-in starts.</span></span>

### <a name="initialize-the-excel-document-with-data"></a><span data-ttu-id="67a38-212">使用数据Excel文档</span><span class="sxs-lookup"><span data-stu-id="67a38-212">Initialize the Excel document with data</span></span>

<span data-ttu-id="67a38-213">通常，当客户从Excel打开一个文档时，他们希望该文档包含网站中的一些数据。</span><span class="sxs-lookup"><span data-stu-id="67a38-213">Typically, when the customer opens up an Excel document from your web site, they expect the document to contain some data from the web site.</span></span> <span data-ttu-id="67a38-214">有两种方法将数据写入文档。</span><span class="sxs-lookup"><span data-stu-id="67a38-214">There are a couple of ways to write data into the document.</span></span>

- <span data-ttu-id="67a38-215">**使用 OOXML SDK 写入数据**。</span><span class="sxs-lookup"><span data-stu-id="67a38-215">**Use the OOXML SDK to write the data**.</span></span> <span data-ttu-id="67a38-216">您可以使用 SDK 直接将任何数据写入文档。</span><span class="sxs-lookup"><span data-stu-id="67a38-216">You can use the SDK to directly write any data into the document.</span></span> <span data-ttu-id="67a38-217">如果您希望数据在文档打开时可用，此方法非常有用。</span><span class="sxs-lookup"><span data-stu-id="67a38-217">This approach is useful if you want the data to be available the instant the document is opened.</span></span>
- <span data-ttu-id="67a38-218">**将自定义查询属性Office加载项**。</span><span class="sxs-lookup"><span data-stu-id="67a38-218">**Pass a custom query property to your Office Add-in**.</span></span> <span data-ttu-id="67a38-219">生成文档时，会为外接程序嵌入一个Office属性，其中包含检索所有所需数据的查询字符串。</span><span class="sxs-lookup"><span data-stu-id="67a38-219">When you generate the document, you embed a custom property for the Office Add-in that contains a query string that retrieves all the required data.</span></span> <span data-ttu-id="67a38-220">外接程序打开后，它将检索查询、运行查询，并使用 Office JS API 将查询结果插入文档中。</span><span class="sxs-lookup"><span data-stu-id="67a38-220">When your add-in opens, it retrieves the query, runs the query, and uses the Office JS API to insert the result of the query into the document.</span></span>

### <a name="working-with-the-ooxml-sdk"></a><span data-ttu-id="67a38-221">使用 OOXML SDK</span><span class="sxs-lookup"><span data-stu-id="67a38-221">Working with the OOXML SDK</span></span>

<span data-ttu-id="67a38-222">OOXML SDK 基于 .NET。</span><span class="sxs-lookup"><span data-stu-id="67a38-222">The OOXML SDK is based on .NET.</span></span> <span data-ttu-id="67a38-223">如果 Web 应用程序没有 .NET，则需要寻找使用 OOXML 的替代方法。</span><span class="sxs-lookup"><span data-stu-id="67a38-223">If your web application does not .NET, you’ll need to look for an alternative way to work with OOXML.</span></span>

<span data-ttu-id="67a38-224">Open [XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs)提供了 OOXML SDK 的 JavaScript 版本。</span><span class="sxs-lookup"><span data-stu-id="67a38-224">There is a JavaScript version of the OOXML SDK available at [Open XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span></span>

<span data-ttu-id="67a38-225">可以将 OOXML 代码放在 Azure 函数中，以将 .NET 代码与 Web 应用程序的其余部分分开。</span><span class="sxs-lookup"><span data-stu-id="67a38-225">You can place the OOXML code in an Azure function to separate the .NET code from the rest of your web application.</span></span> <span data-ttu-id="67a38-226">然后调用 Azure 函数 (从 Web Excel生成) 文档。</span><span class="sxs-lookup"><span data-stu-id="67a38-226">Then call the Azure function (to generate the Excel document) from your Web application.</span></span> <span data-ttu-id="67a38-227">有关 Azure 函数详细信息，请参阅 [Azure 函数简介](/azure/azure-functions/functions-overview)。</span><span class="sxs-lookup"><span data-stu-id="67a38-227">For more information on Azure functions, see [An introduction to Azure Functions](/azure/azure-functions/functions-overview).</span></span>

### <a name="use-single-sign-on"></a><span data-ttu-id="67a38-228">使用单一登录</span><span class="sxs-lookup"><span data-stu-id="67a38-228">Use single sign-on</span></span>

<span data-ttu-id="67a38-229">为了简化身份验证，我们建议你的外接程序实现单一登录。</span><span class="sxs-lookup"><span data-stu-id="67a38-229">To simplify authentication, we recommend your add-in implements single sign-on.</span></span> <span data-ttu-id="67a38-230">有关详细信息，请参阅为加载项[启用Office登录](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="67a38-230">For more information, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="67a38-231">另请参阅</span><span class="sxs-lookup"><span data-stu-id="67a38-231">See also</span></span>

- [<span data-ttu-id="67a38-232">欢迎使用 Open XML SDK 2.5 for Office</span><span class="sxs-lookup"><span data-stu-id="67a38-232">Welcome to the Open XML SDK 2.5 for Office</span></span>](/office/open-xml/open-xml-sdk)
- [<span data-ttu-id="67a38-233">随文档自动打开任务窗格</span><span class="sxs-lookup"><span data-stu-id="67a38-233">Automatically open a task pane with a document</span></span>](../develop/automatically-open-a-task-pane-with-a-document.md)
- [<span data-ttu-id="67a38-234">保留加载项状态和设置</span><span class="sxs-lookup"><span data-stu-id="67a38-234">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="67a38-235">通过提供文件名创建电子表格文档</span><span class="sxs-lookup"><span data-stu-id="67a38-235">Create a spreadsheet document by providing a file name</span></span>](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)