---
title: 从网页打开 Excel 并嵌入 Office 加载项
description: 从网页打开 Excel 并嵌入 Office 加载项。
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: a88cc647fc1dba8ab6e6ddc0b504aab96517026a
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839864"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>从网页打开 Excel 并嵌入 Office 加载项

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="网页上的 Excel 按钮图像，可打开一个新的 Excel 文档，并嵌入加载项并自动打开。":::

扩展 SaaS Web 应用程序，以便客户可以直接从网页打开其数据到 Microsoft Excel。 常见方案是客户将处理 Web 应用程序中的数据。 然后，他们将希望将数据复制到 Excel 文档中。 例如，他们可能需要使用 Excel 执行其他分析。 通常，客户需要将数据导出到文件（如 .csv 文件）中，然后将该数据导入 Excel。 他们还必须手动将 Office 外接程序添加到文档。

将步骤数减少为在生成并打开 Excel 文档的网页上单击一次按钮。 还可以在文档中嵌入 Office 外接程序，在文档打开时显示它。 这可确保客户仍可访问应用程序功能。 当文档打开时，客户选择的数据和 Office 外接程序已可供他们继续工作。

本文介绍了在你自己的 SaaS Web 应用程序中实现此方案的代码和技术。

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>创建新的 Excel 文档并嵌入 Office 加载项

首先，我们了解如何从网页创建 Excel 文档，以及如何在文档中嵌入加载项。 [Office OOXML 嵌入外接程序代码](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)示例演示如何将[Script Lab](https://appsource.microsoft.com/product/office/wa104380862)加载项嵌入新的 Office 文档。 虽然该示例适用于任何 Office 文档，但本文仅重点介绍 Excel 电子表格。 使用以下步骤生成并运行示例。

1. 将示例代码从  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip 计算机中提取到文件夹中。
2. 若要生成并运行示例，请按照自述文件的项目 **部分** 中的步骤操作。
3. 运行示例时，将显示类似于以下屏幕截图的网页。 使用该网页创建一个新的 Excel 文档，该文档在打开时包含 Script Lab。
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="嵌入脚本实验室示例显示的用于选择 Excel 文件并将脚本实验室加载项嵌入其中的网页的屏幕截图。":::

### <a name="how-the-sample-works"></a>示例的工作原理

示例代码使用 OOXML SDK 将 Script Lab 加载项嵌入到您选择的 Excel 文档中。 以下信息取自述文件中"关于[代码"](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md)部分。

文件 **Home.aspx.cs：**

- 提供按钮事件处理程序和基本 UI 操作。
- 使用标准ASP.NET技术上载和下载文件。
- 使用 xlsx、docx 或 pptx (的上载文件的文件扩展名) 确定文件类型。 需要从一开始就完成此操作，因为 Open XML SDK 通常具有每种类型的文件不同的 API。
- 调用 **OOXMLHelper** 以验证文件，并调用 **AddInEmbedder** 以在文件中嵌入 Script Lab，并设置为自动打开。

文件 **AddInEmbedder.cs：**

- 提供主要业务逻辑，此示例中是嵌入 Script Lab 的方法。
- 根据文件类型调用 OOXML 帮助程序。

文件 **OOXMLHelper.cs：**

- 提供所有详细的 OOXML 操作。
- 使用标准技术验证 Office 文件，只需调用 **Document.Open** 方法。 如果文件无效，该方法将引发异常。
- 主要包含由 Open XML 2.5 SDK Productivity Tools 生成的代码，这些工具位于 [Open XML 2.5 SDK 的链接中](/office/open-xml/open-xml-sdk)。

OOXMLHelper.cs中的 **GenerateWebExtensionPart1Content** 方法设置对 Microsoft AppSource 中 Script Lab 的 ID 的引用：

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- **StoreType** 值为"OMEX"，它是 Microsoft AppSource 的别名。
- Store 值为"en-US"，位于 Script Lab 的 Microsoft AppSource 区域性部分。
- **Id** 值是 Script Lab 的 Microsoft AppSource 资产 ID。

如果要从文件共享目录设置外接程序以自动打开，你将使用不同的值：

**StoreType** 值为"FileSystem"。

- **Store** 值是网络共享 URL;例如 \\ \\ ，"MyComputer \\ MySharedFolder"。 这应该是在 Office 信任中心中显示为共享受信任目录地址的确切 URL。
- **Id** 值是加载项清单中的应用程序 ID。
> [!NOTE]
> 有关这些属性的可选值的详细信息，请参阅"[自动打开包含文档的任务窗格"。](../develop/automatically-open-a-task-pane-with-a-document.md)

## <a name="use-the-fluent-ui"></a>使用 Fluent UI

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Word、Excel 和 PowerPoint 的 Fluent UI 图标。":::

最佳做法是使用 Fluent UI 来帮助用户在 Microsoft 产品之间转换。 应始终使用 Office 图标来指示从网页中启动的 Office 应用程序。 让我们修改示例代码以使用 Excel 图标指示它启动 Excel 应用程序。

1. 在 Visual Studio 中打开该Visual Studio。
1. 打开 **Home.aspx** 页。
1. 查找以下代码，该代码是表单上的下载按钮：
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. 将按钮代码替换为以下图像标记。
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. 按 **F5** (**或调试>开始调试) 。** 加载主页时，你将看到图标显示。

有关详细信息，请参阅 Fluent UI 开发人员门户上的 [Office 品牌](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) 图标。  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>将 Excel 文档上载到 Microsoft OneDrive

如果你的客户使用 OneDrive，我们建议将新文档上传到 OneDrive。 这使用户更易于查找和处理文档。 让我们创建新的代码示例，并了解如何使用 Microsoft Graph SDK 将新的 Excel 文档上载到 OneDrive。

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>使用快速入门构建新的 Microsoft Graph Web 应用程序

1. 转到 [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 并按照步骤创建并打开与 Office 365 服务交互的快速启动代码示例。
1. 在 **步骤 1：选择语言或平台，ASP.NET** **MVC。** 虽然此过程中的步骤使用 ASP.NET MVC 选项，但步骤遵循适用于任何语言或平台的模式。
1. 在 **步骤 2：获取应用 ID 和密码**，选择 **"获取应用 ID 和密码"。**
1. 登录到你的 Microsoft 365 帐户。  
1. 在 **"请保存应用密码** "网页上，将应用密码保存到文件位置，稍后可在其中检索和使用。
1. Choose **Got it， take me back to the quick start.**
1. 在 **步骤 2 中：注册成功！** 输入生成的应用密码。
1. 在 **步骤 3：开始编码** 中，选择 **"下载基于 SDK 的代码示例"。**
1. 将下载 zip 文件夹解压缩到本地文件夹。  
1. 在 2019 年 6 月打开 graph-tutorial.sln Visual Studio文件。
1. 生成并运行解决方案，并确认它正常工作。 你应该能够使用日历网页来查看 Microsoft 365 日历。

### <a name="upload-a-file-to-onedrive"></a>将文件上传到 OneDrive

1. 在 Visual Studio 2019 中打开 **graph-tutorial.sln** 解决方案，PrivateSettings.config **文件。**
1. 将新的作用域 **Files.ReadWrite**   添加到 **ida：AppScopes** 项，以便它类似于以下代码：
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. 打开 **Index.cshtml** 文件。
1. 插入以下 ActionLink 代码以创建将文件上传到 OneDrive 的按钮。
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
1. 打开GraphHelper.cs **文件** 。
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
1. 按 **F5** (**或调试>开始调试) 。** Web 应用程序将启动。
1. 选择 **"单击此处登录"，** 然后登录。
1. 选择 **"单击此处"在 OneDrive 上创建新文件**。
1. 打开新的浏览器选项卡并登录到 OneDrive 帐户。 你将在根文件夹中test.txt文件。

现在，你已了解如何将文件上传到 OneDrive，你可以重复使用此代码来上载你创建的任何 Excel 文档。

## <a name="additional-considerations-for-your-solution"></a>解决方案的其他注意事项

每个人的解决方案在技术和方法方面是不同的。 以下注意事项将帮助您规划如何修改解决方案以打开文档和嵌入 Office 外接程序。

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>从网页创建新的 Excel 电子表格

该示例修改现有的 Excel 文档。 更常见的方案是，从网页创建新的 Excel 电子表格。 可以通过提供文件名来查找有关如何在"创建电子表格文档"中创建新电子表格的其他详细信息。 本文演示如何在本地创建文件，但您也可以使用 SpreadsheetDocument.Create 方法上的重载在流中创建文件。

### <a name="read-custom-properties-when-your-add-in-starts"></a>在加载项启动时读取自定义属性

该代码示例使用 OOXML SDK 将代码段 ID 存储在新的 Excel 文档中。 Script Lab 从 Excel 文档中读取代码段 ID，然后在代码段打开时显示该代码段。 您可能需要向自己的外接程序 (（如查询字符串或临时身份验证令牌）发送自定义属性。) 请参阅"保留外接程序状态和设置"，了解有关外接程序启动时如何读取自定义属性的完整详细信息。

### <a name="initialize-the-excel-document-with-data"></a>使用数据初始化 Excel 文档

通常，当客户从您的网站打开 Excel 文档时，他们希望该文档包含网站中的一些数据。 有多种方式将数据写入文档。

- **使用 OOXML SDK 写入数据**。 可以使用 SDK 直接将任何数据写入文档。 如果您希望数据在文档打开时可用，此方法非常有用。
- **将自定义查询属性传递到 Office 外接程序**。 生成文档时，将嵌入 Office 外接程序的自定义属性，该属性包含检索所有所需数据的查询字符串。 加载项打开后，它将检索查询、运行查询，并使用 Office JS API 将查询结果插入文档中。

### <a name="working-with-the-ooxml-sdk"></a>使用 OOXML SDK

OOXML SDK 基于 .NET。 如果 Web 应用程序不是 .NET，则需要寻找使用 OOXML 的替代方法。

在 Open XML SDK for JavaScript 中提供了 OOXML [SDK 的 JavaScript 版本](https://archive.codeplex.com/?p=openxmlsdkjs)。

可以将 OOXML 代码放在 Azure 函数中，以将 .NET 代码与 Web 应用程序的其余部分分开。 然后调用 Azure 函数 (从 Web 应用程序) Excel 文档。 有关 Azure 函数详细信息，请参阅 [Azure 函数简介](/azure/azure-functions/functions-overview)。

### <a name="use-single-sign-on"></a>使用单一登录

为了简化身份验证，我们建议加载项实现单一登录。 有关详细信息，请参阅" [为 Office 加载项启用单一登录"](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>另请参阅

- [欢迎使用 Open XML SDK 2.5 for Office](/office/open-xml/open-xml-sdk)
- [随文档自动打开任务窗格](../develop/automatically-open-a-task-pane-with-a-document.md)
- [保留加载项状态和设置](../develop/persisting-add-in-state-and-settings.md)
- [通过提供文件名创建电子表格文档](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)