---
ms.date: 02/09/2021
ms.prod: non-product-specific
description: 有关如何在 VSTO 加载项与 Office 加载项之间共享代码的教程。
title: 教程：使用共享代码库在 VSTO 加载项与 Office 加载项之间共享代码
ms.localizationpriority: high
ms.openlocfilehash: 76b9e49adcf5954f50794aaae2fdf740c436c480
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711250"
---
# <a name="tutorial-share-code-between-both-a-vsto-add-in-and-an-office-add-in-with-a-shared-code-library"></a>教程：使用共享代码库在 VSTO 加载项与 Office 加载项之间共享代码

Visual Studio Tools for Office (VSTO) 加载项非常适合用于扩展 Office，从而为你的企业或其他企业提供解决方案。 这些加载项已问世很长时间，并且已使用 VSTO 构建上千种解决方案。 但是，它们仅在 Windows 版的 Office 中运行。 无法在 Mac、网页和移动平台上运行 VSTO 加载项。

Office 加载项使用 HTML、JavaScript 和其他 Web 技术来构建所有平台上的 Office 解决方案。 一种好方法是将现有 VSTO 加载项迁移到 Office 加载项，使你的解决方案在所有平台中可用。

你可能想要同时保留具有相同功能的 VSTO 加载项和新 Office 加载项。 这样就能继续为 Windows 版 Office 中使用 VSTO 加载项的客户提供服务。 此外，还能为所有平台的客户提供相同的 Office 加载项功能。 你还可以 [使 Office 加载项与现有 VSTO 加载项兼容](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)。

但是，最好避免为 Office 加载项重写 VSTO 加载项的所有代码。 本教程介绍如何使用这两个加载项的共享代码库来避免重写代码。

## <a name="shared-code-library"></a>共享代码库

本教程将指导你完成在 VSTO 加载项和新式 Office 加载项之间确定和共享通用代码的步骤。 本教程使用非常简单的 VSTO 加载项示例来演示这些步骤，以便你可以专注于在处理自己的 VSTO 加载项时所需的技能和方法。

下图显示了如何将共享代码库用于迁移。 通用代码将重构到新的共享代码库中。 该代码可保持其原始编写语言（例如 C# 或 VB）。 这意味着你可以创建项目引用，从而继续在现有 VSTO 加载项中使用该代码。 创建 Office 加载项时，该加载项也将使用共享代码库，即通过 REST API 对其进行调用。

![使用共享代码库的 VSTO 加载项和 Office 加载项的关系图。](../images/vsto-migration-shared-code-library.png)

本教程中的技能和方法：

- 将代码重构到 .NET 类库中，从而创建共享类库。
- 使用 ASP.NET Core 为共享类库创建 REST API 包装器。
- 从 Office 加载项调用 REST API 来访问共享代码。

## <a name="prerequisites"></a>先决条件

设置开发环境：

1. 安装 [Visual Studio 2019](https://visualstudio.microsoft.com/downloads/)。
1. 安装以下工作负载。
    - ASP.NET 和 Web 开发
    - .NET Core 跨平台开发。
    - Office/SharePoint 开发
    - 以下 **各个** 组件。
        - Visual Studio Tools for Office (VSTO)。
        - .NET Core 3.0 Runtime。

还需要：

- Microsoft 365帐户。你可以加入 [Microsoft 365开发人员计划](https://aka.ms/devprogramsignup)，该计划提供包含 Office 应用的可续订的 90 天 Microsoft 365 订阅。
- Microsoft Azure租户。可在此处获取试用订阅: [Microsoft Azure](https://account.windowsazure.com/SignUp)。

## <a name="the-cell-analyzer-vsto-add-in"></a>单元格分析器 VSTO 加载项

本教程使用 [Office 加载项的 VSTO 加载项共享库](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/VSTO-shared-code-migration) PnP 解决方案。 **/start** 文件夹包含要迁移的 VSTO 加载项解决方案。 你的目标是在可能情况下，通过共享代码，将 VSTO 加载项迁移到新式 Office 加载项。

> [!NOTE]
> 该示例使用 C#，但你可以将本教程中的方法应用于采用任何 .NET 语言编写的 VSTO 加载项。

1. 将 [Office 加载项的 VSTO 加载项共享库](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/VSTO-shared-code-migration) PnP 解决方案下载到计算机上的工作文件夹。
1. 启动 Visual Studio 2019 并打开 **/start/Cell-Analyzer.sln** 解决方案。
1. 在“**调试**”菜单中，选择“**开始调试**”。
1. 在“**解决方案资源管理器**”中，右键单击“**单元格分析器**”项目，然后选择“**属性**”。
1. 在属性中选择“**签名**”类别。
1. 选择“**为 ClickOnce 清单签名**”，然后选择“**创建测试证书**”。
1. 在 **创建测试证书** 对话框中，输入并确认密码。然后选择 **确定**。

该加载项是 Excel 的自定义任务窗格。 你可以选择包含文本的任何单元格，然后选择“**显示 unicode**”按钮。 在“**结果**”部分中，该加载项将列出文本中的每个字符及其相应 Unicode 编号。

![在 Excel 中运行的 Cell Analyzer VSTO 加载项的屏幕截图，带有“显示 unicode”按钮和空结果部分。](../images/pnp-cell-analyzer-vsto-add-in.png)

## <a name="analyze-types-of-code-in-the-vsto-add-in"></a>分析 VSTO 加载项中的代码类型

采用的第一种方法是分析加载项，从而了解可共享代码的哪些部分。 通常情况下，项目将分为三种类型的代码。

### <a name="ui-code"></a>UI 代码

UI 代码与用户进行交互。 在 VSTO 中，UI 代码可通过 Windows 窗体运行。 Office 加载项将 HTML、CSS 和 JavaScript 用于 UI。 由于这些差异，无法将 UI 代码共享到 Office 加载项。 需要用 JavaScript 来重新创建 UI。

### <a name="document-code"></a>文档代码

在 VSTO 中，代码通过 .NET 对象（例如 `Microsoft.Office.Interop.Excel.Range`）与文档进行交互。 但 Office 加载项使用的是 Office.js 库。 虽然它们类似，但是并不完全相同。 同样，不能将文档交互代码共享到 Office 加载项。

### <a name="logic-code"></a>逻辑代码

业务逻辑、算法、helper 函数和类似的代码通常构成 VSTO 加载项的核心。 此类代码独立于 UI 代码和文档代码，可用于执行分析、连接到后端服务、运行计算等。 这是可以共享的代码，因此无需用 JavaScript 重写。

让我们看一看 VSTO 加载项。在以下代码中，每个部分标识为 DOCUMENT、UI 或 ALGORITHM 代码。

```csharp
// *** UI CODE ***
private void btnUnicode_Click(object sender, EventArgs e)
{
    // *** DOCUMENT CODE ***
    Microsoft.Office.Interop.Excel.Range rangeCell;
    rangeCell = Globals.ThisAddIn.Application.ActiveCell;

    string cellValue = "";

    if (null != rangeCell.Value)
    {
        cellValue = rangeCell.Value.ToString();
    }

    // *** ALGORITHM CODE ***
    //convert string to Unicode listing
    string result = "";
    foreach (char c in cellValue)
    {
        int unicode = c;

        result += $"{c}: {unicode}\r\n";
    }

    // *** UI CODE ***
    //Output the result
    txtResult.Text = result;
}
```

使用此方法就会发现，可以将一节代码共享到 Office 加载项。 需要将以下代码重构到单独的类库中。

```csharp
// *** ALGORITHM CODE ***
//convert string to Unicode listing
string result = "";
foreach (char c in cellValue)
{
    int unicode = c;

    result += $"{c}: {unicode}\r\n";
}
```

## <a name="create-a-shared-class-library"></a>创建共享类库

在 Visual Studio 中，VSTO 加载项会创建为 .NET 项目，因此为简单起见，我们将尽可能重用 .NET。 下一种方法是创建类库，然后将共享代码重构到该类库中。

1. 如果尚未启动 Visual Studio 2019 并打开 **\start\Cell-Analyzer.sln** 解决方案，请执行此操作。
1. 右键单击“**解决方案资源管理器**”中的解决方案，并选择 **“添加”>“新建项目”**。
1. 在“**添加新项目**”对话框中，选择“**类库(.NET Framework)**”，然后选择“**下一步**”。
    > [!NOTE]
    > 请勿使用 .NET Core 类库，因为它不能用于你的 VSTO 项目。
1. 在“**配置新项目**”对话框中，设置以下字段。
    - 将“**项目名称**”设置为“**CellAnalyzerSharedLibrary**”。
    - 保留“**位置**”的默认值。
    - 将“**框架**”设置为“**4.7.2**”。
1. 选择“**创建**”。
1. 创建项目后，将 **Class1.cs** 文件重命名为 **CellOperations.cs**。 系统将提示你重命名该类。 请重命名该类名，使其与文件名匹配。
1. 将以下代码添加到 `CellOperations` 类，从而创建名为 `GetUnicodeFromText` 的方法。

    ```csharp
    public class CellOperations
    {
        static public string GetUnicodeFromText(string value)
        {
            string result = "";
            foreach (char c in value)
            {
                int unicode = c;
    
                result += $"{c}: {unicode}\r\n";
            }
            return result;
        }
    }
    ```

### <a name="use-the-shared-class-library-in-the-vsto-add-in"></a>使用 VSTO 加载项中的共享类库

现在，需要更新 VSTO 加载项以使用该类库。 必须确保 VSTO 加载项和 Office 加载项使用同一共享类库，以便在同一位置创建将来的 bug 修复或功能。

1. 在“**解决方案资源管理器**”中，右键单击“**Cell-Analyzer**”项目，然后选择“**添加引用**”。
1. 选择“**CellAnalyzerSharedLibrary**”，然后选择“**确定**”。
1. 在“**解决方案资源管理器**”中，展开“**单元格分析器**”项目，右键单击“**CellAnalyzerPane.cs**”文件，然后选择“**查看代码**”。
1. 在 `btnUnicode_Click` 方法中，删除以下代码行。

    ```csharp
    //Convert to Unicode listing
    string result = "";
    foreach (char c in cellValue)
    {
      int unicode = c;
      result += $"{c}: {unicode}\r\n";
    }
    ```

1. 将 `//Output the result` 注释下的代码行更新为如下代码：

    ```csharp
    //Output the result
    txtResult.Text = CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(cellValue);
    ```

1. 在“**调试**”菜单中，选择“**开始调试**”。 自定义任务窗格应按预期运行。 在单元格中输入一些文本，然后进行测试，以确定可以用加载项将其转换为 Unicode 列表。

## <a name="create-a-rest-api-wrapper"></a>创建 REST API 包装器

VSTO 加载项可以直接使用共享类库，因为它们都是 .NET 项目。 但是，Office 加载项无法使用 .NET，因为它使用的是 JavaScript。 接下来需要创建 REST API 包装器。 这使 Office 加载项可以调用 REST API，进而将此调用传递到共享类库。

1. 在“**解决方案资源管理器**”中，右键单击“**单元格分析器**”项目，然后选择 **“添加”>“新建项目”**。
1. 在“**添加新项目**”对话框中，选择“**ASP.NET Core Web 应用程序**”，然后选择“**下一步**”。
1. 在“**配置新项目**”对话框中，设置以下字段。
    - 将“**项目名称**”设置为“**CellAnalyzerRESTAPI**”。
    - 在“**位置**”字段中，保留默认值。
1. 选择“**创建**”。
1. 在“**创建新的 ASP.NET Core Web 应用程序**”对话框中，选择“**ASP.NET Core 3.1**”版本，然后在项目列表中选择“**API**”。
1. 将其他所有字段保留为默认值，然后选择“**创建**”按钮。
1. 创建项目后，展开“**解决方案资源管理器**”中的“**CellAnalyzerRESTAPI**”项目。
1. 右键单击“**依赖项**”，然后选择“**添加引用**”。
1. 选择“**CellAnalyzerSharedLibrary**”，然后选择“**确定**”。
1. 右键单击“**控制器**”文件夹，然后选择 **“添加”>“控制器”**。
1. 在“**添加新搭建基架的项目**”对话框中，选择“**API 控制器 - 空**”，然后选择“**添加**”。
1. 在“**添加空的 API 控制器**”对话框中，将该控制器命名为“**AnalyzeUnicodeController**”，然后选择“**添加**”。
1. 打开“**AnalyzeUnicodeController.cs**”文件，然后将以下代码作为方法添加到 `AnalyzeUnicodeController` 类。

    ```csharp
    [HttpGet]
    public ActionResult<string> AnalyzeUnicode(string value)
    {
      if (value == null)
      {
        return BadRequest();
      }
      return CellAnalyzerSharedLibrary.CellOperations.GetUnicodeFromText(value);
    }
    ```

1. 右键单击“**CellAnalyzerRESTAPI**”项目，然后选择“**设为启动项目**”。
1. 在“**调试**”菜单中，选择“**开始调试**”。
1. 随后将启动浏览器。 输入以下 URL 以测试 REST API 是否正在运行：`https://localhost:<ssl port number>/api/analyzeunicode?value=test`。 你可以重用 Visual Studio 启动的浏览器中的 URL 的端口号。 应该会看到返回一个字符串，其含有每个字符的 Unicode 值。

## <a name="create-the-office-add-in"></a>创建 Office 加载项

创建 Office 加载项时，将调用 REST API。 但是，首先需要获取 REST API 服务器的端口号并将其保存供以后使用。

### <a name="save-the-ssl-port-number"></a>保存 SSL 端口号

1. 如果尚未启动 Visual Studio 2019 并打开 **\start\Cell-Analyzer.sln** 解决方案，请执行此操作。
1. 在“**CellAnalyzerRESTAPI**”项目中，展开“**属性**”，然后打开 **launchSettings json** 文件。
1. 查找带有 **sslPort** 值的代码行，复制端口号，然后将其保存到某个位置。

### <a name="add-the-office-add-in-project"></a>添加 Office 加载项项目

为简单起见，请将所有代码保存在一个解决方案中。 将 Office 加载项项目添加到现有 Visual Studio 解决方案。 但是，如果你熟悉 [Office 加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md)和 Visual Studio 代码，也可以运行 `yo office` 来生成项目。 操作步骤非常相似。

1. 在“**解决方案资源管理器**”中，右键单击“**单元格分析器**”解决方案，然后选择 **“添加”>“新建项目”**。
1. 在“**添加新项目**”对话框中，选择“**Excel Web 加载项**”，然后选择“**下一步**”。
1. 在“**配置新项目**”对话框中，设置以下字段。
    - 将“**项目名称**”设置为“**CellAnalyzerOfficeAddin**”。
    - 保留“**位置**”的默认值。
    - 将“**框架**”设置为“**4.7.2**”或更高版本。
1. 选择“**创建**”。
1. 在“**选择加载项类型**”对话框中，选择“**将新功能添加到 Excel**”，然后选择“**完成**”。

随后将创建两个项目：

- **CellAnalyzerOfficeAddin** - 此项目将配置用于描述此加载项的清单 XML 文件，以便 Office 可正确将其加载。 其中包含此加载项的 ID、名称、描述和其他信息。
- **CellAnalyzerOfficeAddinWeb** - 此项目包含用于加载项的 Web 资源，如 HTML、CSS 和脚本。 此外，还配置 IIS Express 实例，从而将你的加载项托管为 Web 应用程序。

### <a name="add-ui-and-functionality-to-the-office-add-in"></a>将 UI 和功能添加到 Office 加载项

1. 在“**解决方案资源管理器**”中，展开“**CellAnalyzerOfficeAddinWeb**”项目。
1. 打开 **Home. html** 文件，然后将 `<body>` 的内容替换为以下 HTML。

    ```html
    <button id="btnShowUnicode" onclick="showUnicode()">Show Unicode</button>
    <p>Result:</p>
    <div id="txtResult"></div>
    ```

1. 打开 **Home.js** 文件并将全部内容替换为以下代码。

    ```js
    (function () {
      "use strict";
      // The initialize function must be run each time a new page is loaded.
      Office.initialize = function (reason) {
        $(document).ready(function () {
        });
      };
    })();

    function showUnicode() {
      Excel.run(function (context) {
        const range = context.workbook.getSelectedRange();
        range.load("values");
        return context.sync(range).then(function (range) {
          const url = "https://localhost:<ssl port number>/api/analyzeunicode?value=" + range.values[0][0];
          $.ajax({
            type: "GET",
            url: url,
            success: function (data) {
              let htmlData = data.replace(/\r\n/g, '<br>');
              $("#txtResult").html(htmlData);
            },
            error: function (data) {
                $("#txtResult").html("error occurred in ajax call.");
            }
          });
        });
      });
    }
    ```

1. 在上面的代码中，输入你先前从 **launchSettings json** 文件保存的 **sslPort** 号。

在上面的代码中将处理返回的字符串，以便将回车换行替换为 `<br>` HTML 标记。 偶尔在某些情况下，可能需要在 Office 加载项侧调整 VSTO 加载项中对 .NET 完全适用的返回值，才能使该返回值符合预期。 在此情况下，REST API 和共享类库仅关注返回字符串。 `showUnicode()` 方法负责正确设置返回值的格式以便显示。

### <a name="allow-cors-from-the-office-add-in"></a>允许来自 Office 加载项的 CORS

Office.js 库要求对传出调用执行 CORS，例如 REST API 服务器的 `ajax` 调用便是如此。 若要允许从 Office 加载项调用 REST API，请执行以下步骤。

1. 在“**解决方案资源管理器**”中，选择“**CellAnalyzerOfficeAddinWeb**”项目。
1. 从“**视图**”菜单，选择“**属性窗口**”（如果尚未显示该窗口）。
1. 在属性窗口中，复制“**SSL URL**”的值，然后将其保存到某个位置。 这是你需要通过 CORS 允许的 URL。
1. 在“**CellAnalyzerRESTAPI**”项目中，打开 **Startup.cs** 文件。
1. 将以下代码添加到 `ConfigureServices` 方法上方。 请务必替换先前为 `builder.WithOrigins` 调用复制的 URL SSL。

    ```csharp
    services.AddCors(options =>
    {
      options.AddPolicy(MyAllowSpecificOrigins,
      builder =>
      {
        builder.WithOrigins("<your URL SSL>")
        .AllowAnyMethod()
        .AllowAnyHeader();
      });
    });
    ```

    > [!NOTE]
    > 在 `builder.WithOrigins`方法中使用 URL 时，请保留末尾的 `/`。 例如，它应该类似于 `https://localhost:44000`。 否则，在运行时将出现 CORS 错误。

1. 将下列字段添加到 `Startup` 类：

    ```csharp
    readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
    ```

1. 将以下代码添加到 `configure` 方法中的 `app.UseEndpoints` 代码行之前。

    ```csharp
    app.UseCors(MyAllowSpecificOrigins);
    ```

完成后，`Startup` 类应类似于以下代码（你的 localhost URL 可能有所不同）。

```csharp
public class Startup
{
  public Startup(IConfiguration configuration)
    {
      Configuration = configuration;
    }

    readonly string MyAllowSpecificOrigins = "_myAllowSpecificOrigins";

    public IConfiguration Configuration { get; }

    // NOTE: The following code configures CORS for the localhost:44397 port.
    // This is for development purposes. In production code you should update this to 
    // use the appropriate allowed domains.
    public void ConfigureServices(IServiceCollection services)
    {
        services.AddCors(options =>
        {
            options.AddPolicy(MyAllowSpecificOrigins,
            builder =>
            {
                builder.WithOrigins("https://localhost:44397")
                .AllowAnyMethod()
                .AllowAnyHeader();
            });
        });
        services.AddControllers();
    }

    // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        if (env.IsDevelopment())
        {
            app.UseDeveloperExceptionPage();
        }

        app.UseHttpsRedirection();

        app.UseRouting();

        app.UseAuthorization();

        app.UseCors(MyAllowSpecificOrigins);

        app.UseEndpoints(endpoints =>
        {
            endpoints.MapControllers();
        });
    }
}
```

### <a name="run-the-add-in"></a>运行加载项

1. 在“**解决方案资源管理器**”中，右键单击顶层节点“**解决方案‘单元格分析器’**”，然后选择“**设置启动项目**”。
1. 在“**解决方案‘单元格分析器’属性页**”对话框中，选择“**多个启动项目**”。
1. 为以下每个项目，将“**操作**”属性设置为“**启动**”。

    - CellAnalyzerRESTAPI
    - CellAnalyzerOfficeAddin
    - CellAnalyzerOfficeAddinWeb

1. 选择“**确定**”。
1. 从“**调试**”菜单中选择“**开始调试**”。

Excel 将运行并旁加载 Office 加载项。 可通过以下方式测试 localhost REST API 服务是否正常工作：将文本值输入单元格，然后在 Office 加载项中选择“**显示 Unicode**”按钮。 它应该会调用 REST API 并显示文本字符的 Unicode 值。

## <a name="publish-to-an-azure-app-service"></a>发布到 Azure 应用服务

最终希望将 REST API 项目发布到云。 在以下步骤中，你将了解如何将 **CellAnalyzerRESTAPI** 项目发布到 Microsoft Azure 应用服务。 有关如何获取 Azure 帐户的信息，请参阅“[先决条件](#prerequisites)”。

1. 在“**解决方案资源管理器**”中，右键单击“**CellAnalyzerRESTAPI**”项目，然后选择“**发布**”。
1. 在“**选取发布目标**”对话框中，选择“**新建**”，然后选择“**创建配置文件**”。
1. 在“**应用服务**”对话框中，选择正确帐户（如果尚未选择）。
1. 对于你的帐户，“**应用服务**”对话框的字段将设置为默认值。 通常情况下，默认值可运行良好，但是如果你更希望使用其他设置，则可以更改默认值。
1. 在“**应用服务**”对话框中，选择“**创建**”。
1. 新配置文件将显示在“**发布**”页中。 选择“**发布**”以生成代码并将其部署到应用服务。

现在可以测试该服务。 打开浏览器，输入 URL 直接转至新服务。 例如，使用 `https://<myappservice>.azurewebsites.net/api/analyzeunicode?value=test`，其中 *myappservice* 是你为新应用服务创建的唯一名称。

### <a name="use-the-azure-app-service-from-the-office-add-in"></a>在 Office 加载项中使用 Azure 应用服务

最后一步是更新 Office 加载项中的代码，从而使用 Azure 应用服务，而不是 localhost。

1. 在“**解决方案资源管理器**”中，展开“**CellAnalyzerOfficeAddinWeb**”项目，然后打开“**Home.js**”文件。
1. 将 `url` 常量更改为使用你的 Azure 应用服务的 URL，如以下代码行所示。 将 `<myappservice>` 替换成你为新应用服务创建的唯一名称。

    ```JavaScript
    const url = "https://<myappservice>.azurewebsites.net/api/analyzeunicode?value=" + range.values[0][0];
    ```

1. 在“**解决方案资源管理器**”中，右键单击顶层节点“**解决方案‘单元格分析器’**”，然后选择“**设置启动项目**”。
1. 在“**解决方案‘单元格分析器’属性页**”对话框中，选择“**多个启动项目**”。
1. 为以下每个项目启用“**启动**”操作。
    - CellAnalyzerOfficeAddinWeb
    - CellAnalyzerOfficeAddin
1. 选择“**确定**”。
1. 从“**调试**”菜单中选择“**开始调试**”。

Excel 将运行并旁加载 Office 加载项。 若要测试应用服务是否正常运行，请将文本值输入单元格，然后在 Office 加载项中选择“**显示 Unicode**”。 它应该会调用该服务并显示文本字符的 Unicode 值。

## <a name="conclusion"></a>总结

在本教程中，你学习了如何创建与原始 VSTO 加载项共享代码的 Office 加载项。 学习了如何维护 Windows 版 Office 的 VSTO 代码以及其他平台上的 Office 的 Office 加载项。 你将 VSTO C# 代码重构到共享库中，并将其部署到 Azure 应用服务。 你创建了使用共享库的 Office 加载项，因此无需用 JavaScript 重写代码。
