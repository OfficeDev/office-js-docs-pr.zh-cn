---
title: 从网页中打开 Excel 并嵌入 Office 外接程序
description: 从网页中打开 Excel 并嵌入 Office 外接程序。
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: 49df253c714f3ad84d2523b87e7df894b9027355
ms.sourcegitcommit: ea03e4ea2e8537d5f6d52477816209f6c1a6579c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/21/2020
ms.locfileid: "48166917"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>从网页中打开 Excel 并嵌入 Office 外接程序

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="网页上的 Excel 按钮的图像使用外接程序打开新的 Excel 文档，并将其置于嵌入式和自动打开状态。":::

扩展 SaaS web 应用程序，以便客户可以直接将其数据从网页中直接打开到 Microsoft Excel。 一个常见的情况是，客户将使用 web 应用程序中的数据。 然后，他们需要将数据复制到 Excel 文档中。 例如，他们可能想要使用 Excel 执行其他分析。 通常情况下，客户需要将数据导出到文件（如 .csv 文件），然后将该数据导入到 Excel 中。 他们还必须手动将 Office 加载项添加到文档中。

将步骤数减少为单个按钮单击可生成并打开 Excel 文档的网页。 您还可以将 Office 外接程序嵌入文档中并在文档打开时显示它。 这可确保客户仍有权访问应用程序功能。 当文档打开时，客户选择的数据和你的 Office 外接程序已可供他们继续工作。

本文介绍在您自己的 SaaS web 应用程序中实现此方案的代码和技术。

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>创建新的 Excel 文档并嵌入 Office 加载项

首先，我们来了解如何从网页创建 Excel 文档，并将外接程序嵌入文档中。 [OFFICE OOXML 嵌入加载项代码示例](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)显示了如何将[脚本实验室加载项](https://appsource.microsoft.com/product/office/wa104380862)嵌入到新的 Office 文档中。 虽然本示例适用于任何 Office 文档，但我们只是在本文中重点介绍 Excel 电子表格。 使用以下步骤生成并运行示例。

1. 将示例代码从  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip 您的计算机上的文件夹中提取出来。
2. 若要生成并运行示例，请按照自述文件的 " **使用项目"** 部分中的步骤操作。
3. 运行示例时，它将显示一个类似于以下屏幕截图的网页。 使用网页创建一个在打开时包含脚本实验室的新 Excel 文档。
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="嵌入脚本实验室示例显示用于选择 Excel 文件并将脚本实验室外接程序嵌入到其中的网页的屏幕截图。":::

### <a name="how-the-sample-works"></a>示例的工作原理

示例代码使用 OOXML SDK 将脚本实验室外接程序嵌入到您选择的 Excel 文档中。 以下信息取自自述文件中的 " [**代码"** 部分](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) 。

文件 **Home.aspx.cs**：

- 提供按钮事件处理程序和基本 UI 操作。
- 使用标准 ASP.NET 技术上载和下载文件。
- 使用上传的文件的文件扩展名 (.xlsx、.docx 或 .pptx) 来确定文件的类型。 需要在开始时执行此操作，因为 Open XML SDK 通常对每种类型的文件都具有不同的 Api。
- 调用 **OOXMLHelper** 以验证文件并调用 **AddInEmbedder** 以在文件中嵌入脚本实验室并将其设置为自动打开。

文件 **AddInEmbedder.cs**：

- 提供主要业务逻辑，在此示例中，是一种嵌入脚本实验室的方法。
- 根据文件的类型，对 OOXML 帮助程序进行调用。

文件 **OOXMLHelper.cs**：

- 提供所有详细的 OOXML 操作。
- 使用用于验证 Office 文件的标准技术，这只是调用 **文档的 Open** 方法。 如果文件无效，则该方法将引发异常。
- 主要包含由 Open XML 2.5 SDK 生产力工具生成的代码，这些工具在 [OPEN xml 2.5 sdk](/office/open-xml/open-xml-sdk)的链接中可用。

**OOXMLHelper.cs**文件中的**GenerateWebExtensionPart1Content**方法将引用设置为 Microsoft AppSource 中的脚本实验室的 ID：

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- **StoreType**值为 "OMEX"，为 Microsoft AppSource 的别名。
- **存储**值为脚本实验室的 Microsoft AppSource 区域性部分中的 "en-us"。
- **Id**值是脚本实验室的 Microsoft APPSOURCE 资产 Id。

如果要从文件共享目录中设置自动打开的外接程序，您将使用不同的值：

**StoreType**值为 "FileSystem"。

- **存储**值是网络共享的 URL;例如，" \\ \\ MyComputer \\ MySharedFolder"。 这应该是在 Office 信任中心中显示为共享的受信任目录地址的确切 URL。
- **Id**值是加载项清单中的应用程序 Id。
> [!NOTE]
> 有关这些属性的可选值的详细信息，请参阅 [自动打开包含文档的任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)。

## <a name="use-the-fluent-ui"></a>使用熟知的 UI

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="适用于 Word、Excel 和 PowerPoint 的熟知 UI 图标。":::

最佳做法是使用熟知的 UI 来帮助用户在 Microsoft 产品之间进行转换。 应始终使用 Office 图标来指示将从您的网页启动哪个 Office 应用程序。 让我们修改示例代码，以使用 Excel 图标指示它启动 Excel 应用程序。

1. 打开 Visual Studio 中的示例。
1. 打开 " **主页 .aspx** " 页面。
1. 查找以下代码，它是表单上的 "下载" 按钮：
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. 将按钮代码替换为以下图像标记。
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. 按 **F5** (或 **调试 > 启动调试**) 。 加载主页时，您会看到显示的图标。

有关详细信息，请参阅熟知的 UI 开发人员门户上的 [Office 品牌图标](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) 。  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>将 Excel 文档上载到 Microsoft OneDrive

如果客户使用 OneDrive，建议将新文档上载到 OneDrive。 这样，他们就可以更轻松地查找和使用文档。 我们来创建一个新的代码示例，并了解如何使用 Microsoft Graph SDK 将新的 Excel 文档上载到 OneDrive。

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>使用快速启动构建新的 Microsoft Graph web 应用程序

1. 转到 [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 并按照步骤操作，以创建和打开与 Office 365 服务交互的快速入门代码示例。
1. 在 **步骤1：选择 "语言" 或 "平台**" 中，选择 " **ASP.NET MVC**"。 虽然此过程中的步骤使用 ASP.NET MVC 选项，但这些步骤遵循适用于任何语言或平台的模式。
1. 在 " **步骤2：获取应用程序 id 和密码**" 中，选择 " **获取应用 id 和密码**"。
1. 登录到 Microsoft 365 帐户。  
1. 在 " **请保存您的应用程序机密** 网页" 中，将应用程序密码保存到文件位置，稍后可对其进行检索和使用。
1. 选择 **"已收到"，让我回到 "快速入门"**。
1. 在 **第2步：注册成功！** 输入生成的应用密码。
1. 在 " **步骤3：开始编码**" 中，选择 **"下载基于 SDK 的代码" 示例**。
1. 将下载 zip 文件夹解压缩到本地文件夹中。  
1. 在 Visual Studio 2019 中打开 graph-tutorial 文件。
1. 生成并运行解决方案，并确认它是否正常工作。 您应该能够使用 "日历" 网页查看您的 Microsoft 365 日历。

### <a name="upload-a-file-to-onedrive"></a>将文件上传到 OneDrive

1. 打开 Visual Studio 2019 中的 **graph-tutorial** 解决方案，然后打开 **PrivateSettings.config** 文件。
1. 将新的作用域**文件**添加   到**ida： AppScopes**键，使其类似于以下代码：
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. 打开 " **索引 cshtml** " 文件。
1. 插入以下的 Html.actionlink 代码以创建按钮，以将文件上传到 OneDrive。
    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```
1. 打开 **HomeController.cs** 文件。
1. 插入以下代码以处理来自操作链接的请求。
    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```
1. 打开 **GraphHelper.cs** 文件。
1. 插入以下代码以调用 Microsoft Graph API，以在 OneDrive 上创建新文件。
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
1. 按 **F5** (或 **调试 > 启动调试**) 。 Web 应用程序将启动。
1. 选择 **"单击此处登录**"，然后登录。
1. 选择 " **单击此处可在 OneDrive 上创建新文件"**。
1. 打开一个新的浏览器选项卡并登录到您的 OneDrive 帐户。 您将在根文件夹中看到 test.txt 文件。

现在，您已经了解如何将文件上传到 OneDrive，您可以重复使用此代码来上传您创建的任何 Excel 文档。

## <a name="additional-considerations-for-your-solution"></a>解决方案的其他注意事项

每个人的解决方案在技术和方法方面各不相同。 以下注意事项将帮助您规划如何修改解决方案以打开文档并嵌入 Office 外接程序。

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>从网页创建新 Excel 电子表格

此示例修改现有的 Excel 文档。 一个更常见的方案是，从网页创建一个新的 Excel 电子表格。 您可以通过提供文件名来查找有关如何在 **创建电子表格文档** 中创建新电子表格的其他详细信息。 本文介绍如何在本地创建文件，但您也可以使用 SpreadsheetDocument 方法上的重载在 stream 中创建文件。

### <a name="read-custom-properties-when-your-add-in-starts"></a>在外接程序启动时读取自定义属性

该代码示例使用 OOXML SDK 将一个代码段 ID 存储在新的 Excel 文档中。 脚本实验室从 Excel 文档读取代码段 ID，然后在它打开时显示该代码段。 您可能需要将自定义属性发送到您自己的外接程序 (例如查询字符串或临时身份验证令牌。 ) 请参阅 **保留外接程序状态和设置** ，了解有关如何在加载项启动时读取自定义属性的完整详细信息。

### <a name="initialize-the-excel-document-with-data"></a>使用数据初始化 Excel 文档

通常，当客户从您的网站打开 Excel 文档时，他们希望文档包含网站中的一些数据。 有几种方法可将数据写入文档中。

- **使用 OOXML SDK 写入数据**。 您可以使用 SDK 直接将任何数据写入文档中。 如果您希望数据在文档打开时即时可用，则此方法很有用。
- 将**自定义查询属性传递到 Office 外接程序**。 在生成文档时，您嵌入了 Office 加载项的自定义属性，该属性包含检索所有必需数据的查询字符串。 当您的外接程序打开时，它将检索查询，运行查询，并使用 Office JS API 将查询结果插入到文档中。

### <a name="working-with-the-ooxml-sdk"></a>使用 OOXML SDK

OOXML SDK 基于 .NET。 如果您的 web 应用程序不是 .NET，则需要查找另一种使用 OOXML 的方法。

在适用于 [javascript 的 OPEN XML SDK](https://archive.codeplex.com/?p=openxmlsdkjs)中，有一个适用于 OOXML Sdk 的 JavaScript 版本。

您可以将 OOXML 代码放在 Azure 函数中，以将 .NET 代码与 web 应用程序的其余部分分开。 然后，调用 Azure 函数 (从 Web 应用程序生成 Excel 文档) 。 有关 Azure 函数的详细信息，请参阅 [Azure 函数简介](https://docs.microsoft.com/azure/azure-functions/functions-overview)。

### <a name="simplify-authentication"></a>简化身份验证

通常情况下，在 web 应用程序中工作时，将对客户进行身份验证并登录。 一种最佳做法是允许他们在打开文档时保持登录状态，这样他们就无需再次登录即可使用 Office 外接程序。 处理此问题的一种良好方式是将生存期为的身份验证令牌传递给加载项。

1. 使用 OOXML SDK 将身份验证令牌另存为文档中的自定义属性。
1. 当加载项启动时，从文档中读取标记。
1. 然后，外接程序可以连接到您的服务，而无需从客户进行任何其他身份验证步骤。

> [!WARNING]
> 在文档中嵌入身份验证令牌会带来安全风险，在未经授权的用户可以获取令牌的情况下。 我们建议使用生存期较短的身份验证令牌。 当外接程序使用短寿命令牌时，它应立即请求未保存在文档中的新的身份验证令牌。

## <a name="see-also"></a>另请参阅

- [欢迎使用 Open XML SDK 2.5 for Office](/office/open-xml/open-xml-sdk)
- [随文档自动打开任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)
- [保留加载项状态和设置](../develop/persisting-add-in-state-and-settings.md)
- [通过提供文件名创建电子表格文档](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)
