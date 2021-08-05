---
title: 从Excel打开加载项并嵌入Office加载项
description: 从Excel打开"加载项"，并嵌入Office加载项。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: a7998d1f15f40a549f8ff9ddd9745d6bf9b8ab6d
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773137"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>从Excel打开加载项并嵌入Office加载项

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="网页上Excel按钮的图像，该按钮可打开一个新的 Excel 文档，并嵌入并自动打开外接程序。":::

扩展 SaaS Web 应用程序，以便客户可以直接在网页中打开其数据Microsoft Excel。 一种常见方案是客户将处理 Web 应用程序中的数据。 然后，他们希望将数据复制到一个Excel文档中。 例如，他们可能需要使用数据进行其他Excel。 通常，客户需要将数据导出到文件（如 .csv 文件）中，然后将该数据导入Excel。 他们还必须手动将Office加载项添加到文档中。

将步骤数减少为生成文档并打开文档的网页上的单个Excel单击。 您还可以在文档中Office外接程序，在文档打开时显示它。 这将确保客户仍可访问你的应用程序功能。 当文档打开时，客户选择的数据以及你的Office外接程序已可供他们继续工作。

本文介绍了在你自己的 SaaS Web 应用程序中实现此方案的代码和技术。

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>新建Excel文档并嵌入Office加载项

首先，让我们了解如何从网页Excel文档，以及如何在文档中嵌入加载项。 the [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the Script Lab [add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document. 尽管该示例适用于Office文档，但我们将重点介绍本文Excel电子表格。 使用以下步骤生成并运行示例。

1. 将示例代码从  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip 中提取到您计算机的文件夹中。
2. 若要生成并运行示例，请按照自述文件" **使用项目"** 部分的步骤操作。
3. 运行示例时，将显示类似于以下屏幕截图的网页。 使用网页创建一个新的Excel文档，其中包含Script Lab打开时的内容。
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="嵌入脚本实验室示例显示的网页的屏幕截图，用于选择Excel文件并将脚本实验室外接程序嵌入其中。":::

### <a name="how-the-sample-works"></a>示例的工作原理

示例代码使用 OOXML SDK 将Script Lab嵌入到您Excel文档。 以下信息来自自述 [**文件的关于**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md)代码部分。

文件 **Home.aspx.cs：**

- 提供按钮事件处理程序和基本 UI 操作。
- 使用 ASP.NET 技术上载和下载文件。
- 使用 xlsx、docx 或 pptx (上传的文件的文件扩展名) 确定文件类型。 需要从一开始就完成此操作，因为 Open XML SDK 通常对于每种类型的文件都有不同的 API。
- 调用 **OOXMLHelper** 以验证文件，并调用 **AddInEmbedder** 以在Script Lab嵌入文件并设置为自动打开。

文件 **AddInEmbedder.cs**：

- 提供主要业务逻辑，此示例中是嵌入 Script Lab。
- 根据文件类型调用 OOXML 帮助程序。

文件 **OOXMLHelper.cs：**

- 提供所有详细的 OOXML 操作。
- 使用标准技术来验证Office文件，只需对该文件 **调用 Document.Open** 方法。 如果文件无效，该方法将引发异常。
- 包含主要由 Open XML 2.5 SDK Productivity Tools 生成的代码，这些代码位于 [Open XML 2.5 SDK 的链接中](/office/open-xml/open-xml-sdk)。

**OOXMLHelper.cs** 文件中 **GenerateWebExtensionPart1Content** 方法设置对 Microsoft AppSource 中 Script Lab ID 的引用：

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- **StoreType** 值为"OMEX"，它是 Microsoft AppSource 的别名。
- Store 值为"en-US"，可以在 Microsoft AppSource 区域性部分找到Script Lab。
- **Id** 值是 Microsoft AppSource 资产 ID Script Lab。

如果要从文件共享目录设置外接程序以自动打开，你将使用不同的值：

**StoreType** 值为"FileSystem"。

- **Store** 值是网络共享 URL;例如 \\ \\ ，"MyComputer \\ MySharedFolder"。 这应该是在共享信任中心显示为共享受信任目录地址Office URL。
- **Id** 值是外接程序清单中的应用程序 ID。
> [!NOTE]
> 有关这些属性的可选值的详细信息，请参阅 [使用文档自动打开任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)。

## <a name="use-the-fluent-ui"></a>使用 Fluent UI

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="FluentWord、Excel 和 PowerPoint 的 UI 图标。":::

最佳做法是使用 Fluent UI 来帮助用户在 Microsoft 产品之间过渡。 应始终使用Office图标来指示Office从网页启动哪个应用程序。 让我们修改示例代码，以使用 Excel 图标指示它启动 Excel 应用程序。

1. 在"管理"中Visual Studio。
1. 打开 **Home.aspx** 页。
1. 在表单上查找以下作为下载按钮的代码。

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. 将按钮代码替换为以下图像标记。

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. 按 **F5** (**或调试>开始调试**) 。 加载主页时，你将看到图标出现。

有关详细信息，请参阅[Office](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) UI 开发人员门户Fluent品牌图标。  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Upload Excel文档Microsoft OneDrive

如果你的客户使用 OneDrive，我们建议将新文档OneDrive。 这使用户更易于查找并处理文档。 让我们创建新的代码示例，了解如何使用 Microsoft Graph SDK 将新的 Excel 文档上载到OneDrive。

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>使用快速入门生成新的 Microsoft Graph Web 应用程序

1. 转到 [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 并按照步骤创建并打开与服务交互的快速启动Office示例。
1. 在 **步骤 1：选择语言或平台中**，选择 **"ASP.NET MVC"。** 虽然此过程中的步骤使用 ASP.NET MVC 选项，但步骤遵循适用于任何语言或平台的模式。
1. 在 **步骤 2：获取应用 ID 和密码中，** 选择 **"获取应用 ID 和密码"。**
1. 登录到你的 Microsoft 365 帐户。  
1. 在 **"请保存应用密码** "网页上，将应用密码保存到稍后可以检索和使用的文件位置。
1. 选择 **"已接受"，将我返回到快速入门**。
1. 在 **步骤 2：注册成功！** 输入生成的应用密码。
1. 在 **"步骤 3： 开始编码"中**，**选择"下载基于 SDK 的代码示例"。**
1. 将下载 zip 文件夹解压缩到本地文件夹。  
1. 在 2019 年 10 月Visual Studio graph-tutorial.sln 文件。
1. 生成并运行解决方案并确认它正常工作。 您应该能够使用日历网页来查看您的日历Microsoft 365日历。

### <a name="upload-a-file-to-onedrive"></a>Upload文件OneDrive

1. 在 Visual Studio 2019 中打开 **graph-tutorial.sln** 解决方案，PrivateSettings.config **文件。**

1. 将新的作用域 **Files.ReadWrite**   添加到 **ida：AppScopes** 项，以便它类似于以下代码。

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. 打开 **Index.cshtml** 文件。
1. 插入以下 ActionLink 代码以创建一个按钮以将文件上载到OneDrive。

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
1. 插入以下代码以调用 Microsoft Graph API，以在 OneDrive。

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

1. 按 **F5** (**或调试>开始调试**) 。 Web 应用程序将启动。
1. 选择 **"单击此处登录"，** 然后登录。
1. 选择 **"单击此处"以在"新建OneDrive"。**
1. 打开新的浏览器选项卡，然后登录你的OneDrive帐户。 你将看到根文件夹中test.txt文件。

现在，你已了解如何将文件上载到OneDrive，可以重复使用此代码上载Excel创建的任何文档。

## <a name="additional-considerations-for-your-solution"></a>解决方案的其他注意事项

每个人的解决方案在技术和方法方面是不同的。 以下注意事项将帮助您规划如何修改解决方案以打开文档并嵌入Office外接程序。

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>从网页Excel新建一个电子表格

本示例修改现有文档Excel文档。 更常见的方案是，从网页Excel一个新的电子表格。 在"通过提供文件名创建电子表格文档"中，可以找到 **有关新建电子表格** 的其他详细信息。 本文演示如何在本地创建文件，但您也可以使用 SpreadsheetDocument.Create 方法上的重载在流中创建文件。

### <a name="read-custom-properties-when-your-add-in-starts"></a>在加载项启动时读取自定义属性

该代码示例使用 OOXML SDK 将代码段 ID Excel文档。 Script Lab从文档读取代码Excel ID，然后在代码段打开时显示该代码段。 您可能需要将自定义属性发送到您自己的外接程序 (例如查询字符串或临时身份验证令牌。) 请参阅持久化外接程序状态和设置，了解有关在外接程序启动时如何读取自定义属性的完整详细信息。

### <a name="initialize-the-excel-document-with-data"></a>使用数据Excel文档

通常，当客户从Excel打开一个文档时，他们希望该文档包含网站中的一些数据。 有两种方法将数据写入文档。

- **使用 OOXML SDK 写入数据**。 您可以使用 SDK 直接将任何数据写入文档。 如果您希望数据在文档打开时可用，此方法非常有用。
- **将自定义查询属性Office加载项**。 生成文档时，会为外接程序嵌入一个Office属性，其中包含检索所有所需数据的查询字符串。 外接程序打开后，它将检索查询、运行查询，并使用 Office JS API 将查询结果插入文档中。

### <a name="working-with-the-ooxml-sdk"></a>使用 OOXML SDK

OOXML SDK 基于 .NET。 如果 Web 应用程序没有 .NET，则需要寻找使用 OOXML 的替代方法。

Open [XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs)提供了 OOXML SDK 的 JavaScript 版本。

可以将 OOXML 代码放在 Azure 函数中，以将 .NET 代码与 Web 应用程序的其余部分分开。 然后调用 Azure 函数 (从 Web Excel生成) 文档。 有关 Azure 函数详细信息，请参阅 [Azure 函数简介](/azure/azure-functions/functions-overview)。

### <a name="use-single-sign-on"></a>使用单一登录

为了简化身份验证，我们建议你的外接程序实现单一登录。 有关详细信息，请参阅为加载项[启用Office登录](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>另请参阅

- [欢迎使用 Open XML SDK 2.5 for Office](/office/open-xml/open-xml-sdk)
- [随文档自动打开任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)
- [保留加载项状态和设置](../develop/persisting-add-in-state-and-settings.md)
- [通过提供文件名创建电子表格文档](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)