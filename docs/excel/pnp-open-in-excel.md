---
title: 从网页打开 Excel 并嵌入 Office 加载项
description: 从网页打开 Excel 并嵌入 Office 加载项。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 835518fb822602d6ca1af633f96d2be1e2697f44
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810342"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>从网页打开 Excel 并嵌入 Office 加载项

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="网页上的“Excel”按钮的图像打开一个新的 Excel 文档，其中嵌入了加载项并自动打开。":::

扩展 SaaS Web 应用程序，以便客户可以直接将他们的数据从网页打开到 Microsoft Excel。 一种常见方案是客户将在 Web 应用程序中处理数据。 然后，他们希望将数据复制到 Excel 文档中。 例如，他们可能想要使用 Excel 执行其他分析。 通常，客户需要将数据导出到文件（例如.csv文件），然后将该数据导入 Excel。 他们还必须手动将 Office 加载项添加到文档。

将步骤数减少为在网页上单击一个按钮，以生成并打开 Excel 文档。 还可以在文档中嵌入 Office 加载项，并在文档打开时显示它。 这可确保客户仍有权访问应用程序功能。 文档打开时，客户选择的数据以及 Office 加载项已可供他们继续使用。

本文介绍在你自己的 SaaS Web 应用程序中实现此方案的代码和技术。

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>创建新的 Excel 文档并嵌入 Office 加载项

首先，让我们了解如何从网页创建 Excel 文档，以及如何在文档中嵌入加载项。 [Office OOXML 嵌入外接程序代码示例](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)演示如何将 [Script Lab 外接程序](https://appsource.microsoft.com/product/office/wa104380862)嵌入到新的 Office 文档中。 尽管此示例适用于任何 Office 文档，但本文仅重点介绍 Excel 电子表格。 使用以下步骤生成并运行示例。

1. 将示例代码提取  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip 到计算机上的文件夹中。
2. 若要生成并运行示例，请按照自述文件的 **“使用项目** ”部分中的步骤操作。
3. 运行示例时，将显示类似于以下屏幕截图的网页。 使用网页创建一个新的 Excel 文档，该文档在打开时包含Script Lab。
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="嵌入脚本实验室示例显示的网页的屏幕截图，用于选择 Excel 文件并将脚本实验室加载项嵌入其中。":::

### <a name="how-the-sample-works"></a>示例的工作原理

示例代码使用 OOXML SDK 将Script Lab加载项嵌入到所选的 Excel 文档。 以下信息取自自文件关于 [**代码** 部分](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) 。

文件 **Home.aspx.cs**：

- 提供按钮事件处理程序和基本 UI 操作。
- 使用标准 ASP.NET 技术上传和下载文件。
- 使用上传的文件的文件扩展名 (xlsx、docx 或 pptx) 来确定文件类型。 这需要在一开始完成，因为 Open XML SDK 通常对每种类型的文件具有不同的 API。
- 调用 **OOXMLHelper** 以验证文件，并调用 **AddInEmbedder** 以在文件中嵌入Script Lab并设置为自动打开。

文件 **AddInEmbedder.cs**：

- 提供主业务逻辑，此示例中是嵌入Script Lab的方法。
- 根据文件类型调用 OOXML 帮助程序。

文件 **OOXMLHelper.cs**：

- 提供所有详细的 OOXML 操作。
- 使用标准技术来验证 Office 文件，只需对其调用 **Document.Open** 方法。 如果文件无效，该方法将引发异常。
- 主要包含由 Open XML 2.5 SDK Productivity Tools 生成的代码，这些代码可在 [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk) 的链接中找到。

**OOXMLHelper.cs** 文件中的 **GenerateWebExtensionPart1Content** 方法设置对 Microsoft AppSource 中Script Lab ID 的引用：

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- **StoreType** 值为“OMEX”，这是 Microsoft AppSource 的别名。
- Microsoft AppSource 区域性部分中的 **Microsoft** AppSource 值是“en-US”Script Lab。
- **Id** 值是Script Lab的 Microsoft AppSource 资产 ID。

如果要从文件共享目录设置加载项以自动打开，将使用不同的值：

**StoreType** 值为“FileSystem”。

- **Store** 值是网络共享的 URL;例如，“\\\\MyComputer\\MySharedFolder”。 这应该是在 Office 信任中心显示为共享的受信任目录地址的确切 URL。
- **Id** 值是外接程序清单中的应用 ID。
> [!NOTE]
> 有关这些属性的替代值的详细信息，请参阅 [自动打开包含文档的任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)。

## <a name="use-the-fluent-ui"></a>使用 Fluent UI

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Word、Excel 和 PowerPoint 的 Fluent UI 图标。":::

最佳做法是使用 Fluent UI 来帮助用户在 Microsoft 产品之间切换。 应始终使用 Office 图标来指示将从网页启动哪个 Office 应用程序。 让我们修改示例代码以使用 Excel 图标来指示它启动 Excel 应用程序。

1. 在 Visual Studio 中打开示例。
1. 打开 **“Home.aspx** ”页。
1. 查找以下代码，该代码是窗体上的下载按钮。

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. 将按钮代码替换为以下图像标记。

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. 按 **F5** (或 **调试** > **启动调试**) 。 加载主页时，你将看到图标。

有关详细信息，请参阅 Fluent UI 开发人员门户中的 [Office 品牌图标](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) 。  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>将 Excel 文档上传到 Microsoft OneDrive

如果你的客户使用 OneDrive，我们建议将新文档上传到 OneDrive。 这使得他们更容易查找和使用文档。 让我们创建新的代码示例，并了解如何使用 Microsoft Graph SDK 将新的 Excel 文档上传到 OneDrive。

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>使用快速入门生成新的 Microsoft Graph Web 应用程序

1. 转到 [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 并按照步骤创建并打开与 Office 服务交互的快速入门代码示例。
1. 在 **步骤 1：选择语言或平台** 中，选择 **“ASP.NET MVC**”。 尽管此过程中的步骤使用 ASP.NET MVC 选项，但这些步骤遵循适用于任何语言或平台的模式。
1. 在 **步骤 2：获取应用 ID 和机密** 中，选择 **“获取应用 ID 和机密**”。
1. 登录到 Microsoft 365 帐户。  
1. 在“ **请保存应用机密** ”网页上，将应用机密保存到文件位置，以便稍后检索和使用它。
1. 选择 **“了解”，将我带回快速入门**。
1. **步骤 2：注册成功！** 输入生成的应用机密。
1. 在 **步骤 3：开始编码** 中，选择 **“下载基于 SDK 的代码示例**”。
1. 将下载 zip 文件夹提取到本地文件夹中。  
1. 在 Visual Studio 2019 中打开 graph-tutorial.sln 文件。
1. 生成并运行解决方案，并确认其正常工作。 你应该能够使用日历网页查看 Microsoft 365 日历。

### <a name="upload-a-file-to-onedrive"></a>将文件上传到 OneDrive

1. 在 Visual Studio 2019 中打开 **graph-tutorial.sln** 解决方案，然后打开 **PrivateSettings.config** 文件。

1. 将新的作用域 **Files.ReadWrite** 添加到 **ida：AppScopes** 键，使其类似于以下代码。

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. 打开 **Index.cshtml** 文件。
1. 插入以下 ActionLink 代码以创建按钮以将文件上传到 OneDrive。

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. 打开 **HomeController.cs** 文件。
1. 插入以下代码来处理来自操作链接的请求。

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
1. 插入以下代码以调用 Microsoft Graph API 以在 OneDrive 上创建新文件。

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

1. 按 **F5** (或 **调试** > **启动调试**) 。 Web 应用程序将启动。
1. 选择 **“单击此处登录**”，然后登录。
1. 选择 **“单击此处”，在 OneDrive 上创建新文件**。
1. 打开新的浏览器选项卡并登录到 OneDrive 帐户。 你将在根文件夹中看到 test.txt 文件。

了解如何将文件上传到 OneDrive 后，可以重复使用此代码来上传创建的任何 Excel 文档。

## <a name="additional-considerations-for-your-solution"></a>解决方案的其他注意事项

每个人的解决方案在技术和方法方面都不同。 以下注意事项将帮助你规划如何修改解决方案以打开文档和嵌入 Office 加载项。

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>从网页创建新的 Excel 电子表格

该示例修改现有 Excel 文档。 更常见的方案是从网页创建新的 Excel 电子表格。 有关如何创建新电子表格的其他详细信息，请参阅通过提供文件名 **创建电子表格文档** 。 本文介绍如何在本地创建文件，但你也可以通过在 SpreadsheetDocument.Create 方法上使用重载在流中创建文件。

### <a name="read-custom-properties-when-your-add-in-starts"></a>在加载项启动时读取自定义属性

该代码示例使用 OOXML SDK 将代码片段 ID 存储在新的 Excel 文档中。 Script Lab从 Excel 文档读取代码片段 ID，然后在打开时显示该代码片段代码。 可能需要将自定义属性发送到自己的外接程序 (，例如查询字符串或临时身份验证令牌。) 请参阅 **持久化加载项状态和设置** ，详细了解如何在外接程序启动时读取自定义属性。

### <a name="initialize-the-excel-document-with-data"></a>使用数据初始化 Excel 文档

通常，当客户从您的网站打开 Excel 文档时，他们希望该文档包含来自网站的一些数据。 可通过多种方法将数据写入文档。

- **使用 OOXML SDK 写入数据**。 可以使用 SDK 直接将任何数据写入文档。 如果希望数据在文档打开时立即可用，则此方法非常有用。
- **将自定义查询属性传递给 Office 外接程序**。 生成文档时，将嵌入 Office 外接程序的自定义属性，该属性包含检索所有所需数据的查询字符串。 当加载项打开时，它将检索查询、运行查询，并使用 Office JS API 将查询结果插入文档中。

### <a name="working-with-the-ooxml-sdk"></a>使用 OOXML SDK

OOXML SDK 基于 .NET。 如果 Web 应用程序不是 .NET，则需要寻找使用 OOXML 的替代方法。

可以将 OOXML 代码放在 Azure 函数中，以将 .NET 代码与 Web 应用程序的其余部分分开。 然后调用 Azure 函数 (从 Web 应用程序生成 Excel 文档) 。 有关 Azure Functions 的详细信息，请参阅 [Azure Functions 简介](/azure/azure-functions/functions-overview)。

### <a name="use-single-sign-on"></a>使用单一登录

为了简化身份验证，我们建议加载项实现单一登录。 有关详细信息，请参阅 [为 Office 加载项启用单一登录](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>另请参阅

- [欢迎使用 Open XML SDK 2.5 for Office](/office/open-xml/open-xml-sdk)
- [随文档自动打开任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)
- [保留加载项状态和设置](../develop/persisting-add-in-state-and-settings.md)
- [通过提供文件名创建电子表格文档](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)