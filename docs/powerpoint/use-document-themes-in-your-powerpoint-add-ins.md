# <a name="use-document-themes-in-your-powerpoint-add-ins"></a>在您的 PowerPoint 外接程序中使用文档主题


[Office 主题](https://support.office.com/en-US/Article/What-is-a-theme--7528ccc2-4327-4692-8bf5-9b5a3f2a5ef5) 一部分包括视觉协调的字体和颜色集，可将其应用于演示文稿、文档、工作表以及电子邮件。要在 PowerPoint 中应用或自定义演示文稿的主题，请使用功能区的“**设计**”选项卡上的“**主题**”和“**变量**”组。PowerPoint 通过默认的“**Office 主题**”分配新的空白演示文稿，但也可以选择“**设计**”选项卡上可用的其他主题，从 Office.com 下载其他主题，或创建并自定义自己的主题。

使用 OfficeThemes.css，有助于你以两种方式设计与 PowerPoint 相协调的外接程序：


-  **在 PowerPoint 内容加载项中** 使用 OfficeThemes.css 的文档主题类来指定与内容 外接程序 要插入的演示文稿主题匹配的字体和颜色（这些颜色和字体将在用户更改或自定义演示文稿主题时动态更新）。
    
-  **在 PowerPoint 任务窗格加载项中** 使用 OfficeThemes.css 的 Office UI 主题类来指定 UI 中使用的相同字体和背景颜色，这样您的任务窗格加载项将会与内置任务窗格的颜色匹配（这些颜色将在用户更改 Office UI 主题时动态更新）。
    

### <a name="document-theme-colors"></a>文档主题颜色

每个 Office 文档主题定义了 12 种颜色。在您通过颜色选取器在演示文稿中设置字体、背景和其他颜色设置时，有 10 种颜色可选：


![调色板](../../images/off15app_ColorPalette.png)

若要在 PowerPoint 中查看或自定义一套完整的 12 种主题颜色，请在“**设计**”选项卡的“**变量**”组中，单击“**更多**”下拉菜单，然后指向“**颜色**”，并单击“**自定义颜色**”以显示“**新建主题颜色**”对话框：


![“新建主题颜色”对话框](../../images/off15app_CreateNewThemeColors.png)

前四种颜色适用于文本和背景。使用浅色创建的文本始终在深色背景上清晰显示，使用深色创建的文本始终在浅色背景上清晰显示。接下来六种颜色是个性色，始终在四种潜在背景色上可见。最后两种颜色适用于超链接和已访问过的超链接。


### <a name="document-theme-fonts"></a>文档主题字体

每个 Office 文档主题还定义两种字体 -- 一种用于标题，另一种用于正文文本。PowerPoint 使用这些字体来构造自动文本样式。此外，文本和**艺术字**的**快速样式**库使用这些相同的主题字体。当你使用字体选取器选择字体时，这两种字体就是可用的前两个选项：


![字体选取器](../../images/off15app_FontPicker.png)

若要在 PowerPoint 中查看或自定义主题字体，请在“**设计**”选项卡的“**变量**”组中，单击“**更多**”下拉菜单，然后指向“**字体**”，并单击“**自定义字体**”以显示“**新建主题字体**”对话框。


![“新建主题字体”对话框](../../images/off15app_CreateNewThemeFonts.png)


### <a name="office-ui-theme-fonts-and-colors"></a>Office UI 主题字体和颜色

Office 还让你在几个预定义主题之间进行选择，指定用于所有 Office 应用程序 UI 中的一些颜色和字体。若要执行此操作，请使用“**文件**” > “**帐户**” > “**Office 主题**”下拉菜单（在任意 Office 应用程序中）。


![Office 主题下拉菜单](../../images/off15app_OfficeThemePicker.png)

OfficeThemes.css 包含您可在 PowerPoint 任务窗格加载项中使用的类，以便它们使用这些相同的字体和颜色。这可使您设计与内置任务窗格外观一致的任务窗格加载项。


## <a name="using-officethemescss"></a>使用 OfficeThemes.css

使用 OfficeThemes.css 文件处理 PowerPoint 内容加载项，使您可将 外接程序 的外观与它运行的演示文稿所应用的主题相协调。使用 OfficeThemes.css 文件处理 PowerPoint 任务窗格加载项，使您可将您 外接程序 的外观与 Office UI 的字体和颜色相协调。


### <a name="adding-the-officethemescss-file-to-your-project"></a>将 OfficeThemes.css 文件添加到您的项目中

使用以下步骤将 OfficeThemes.css 文件添加到您的 外接程序 项目中并进行引用。


### <a name="to-add-officethemescss-to-your-visual-studio-project"></a>将 OfficeThemes.css 添加到您的 Visual Studio 项目中


1. 在“**解决方案资源管理器**”中的_“**project_name**”_“**Web**”项目中右键单击“**内容**”文件夹，指向“**添加**”，然后单击“**样式表**”。
    
2. 将新的样式表命名为 OfficeThemes。
    
     >**重要说明**  样式表必须是已命名的 OfficeThemes，或者是当用户更改主题不起作用时可动态更新外接程序字体和颜色的功能。
3. 删除文件中默认的  **body** 类 ( `body {}`)，并将以下 CSS 代码复制并粘贴到文件中。
    
```
  /* The following classes describe the common theme information for office documents */ /* Basic Font and Background Colors for text */ .office-docTheme-primary-fontColor { color:#000000; } .office-docTheme-primary-bgColor { background-color:#ffffff; } .office-docTheme-secondary-fontColor { color: #000000; } .office-docTheme-secondary-bgColor { background-color: #ffffff; } /* Accent color definitions for fonts */ .office-contentAccent1-color { color:#5b9bd5; } .office-contentAccent2-color { color:#ed7d31; } .office-contentAccent3-color { color:#a5a5a5; } .office-contentAccent4-color { color:#ffc000; } .office-contentAccent5-color { color:#4472c4; } .office-contentAccent6-color { color:#70ad47; } /* Accent color for backgrounds */ .office-contentAccent1-bgColor { background-color:#5b9bd5; } .office-contentAccent2-bgColor { background-color:#ed7d31; } .office-contentAccent3-bgColor { background-color:#a5a5a5; } .office-contentAccent4-bgColor { background-color:#ffc000; } .office-contentAccent5-bgColor { background-color:#4472c4; } .office-contentAccent6-bgColor { background-color:#70ad47; } /* Accent color for borders */ .office-contentAccent1-borderColor { border-color:#5b9bd5; } .office-contentAccent2-borderColor { border-color:#ed7d31; } .office-contentAccent3-borderColor { border-color:#a5a5a5; } .office-contentAccent4-borderColor { border-color:#ffc000; } .office-contentAccent5-borderColor { border-color:#4472c4; } .office-contentAccent6-borderColor { border-color:#70ad47; } /* links */ .office-a { color: #0563c1; } .office-a:visited { color: #954f72; } /* Body Fonts */ .office-bodyFont-eastAsian { } /* East Asian name of the Font */ .office-bodyFont-latin { font-family:"Calibri"; } /* Latin name of the Font */ .office-bodyFont-script { } /* Script name of the Font */ .office-bodyFont-localized { font-family:"Calibri"; } /* Localized name of the Font. Corresponds to the default font of the culture currently used in Office.*/ /* Headers Font */ .office-headerFont-eastAsian { } .office-headerFont-latin { font-family:"Calibri Light"; } .office-headerFont-script { } .office-headerFont-localized { font-family:"Calibri Light"; } /* The following classes define font and background colors for Office UI themes. These classes should only be used in task pane add-ins */ /* Basic Font and Background Colors for PPT */ .office-officeTheme-primary-fontColor { color:#b83b1d; } .office-officeTheme-primary-bgColor { background-color:#dedede; } .office-officeTheme-secondary-fontColor { color:#262626; } .office-officeTheme-secondary-bgColor { background-color:#ffffff; } 
```

4. 如果您使用非 Visual Studio 的工具来创建您的 外接程序，请将步骤 3 的 CSS 代码复制到文本文件中，确保将文件保存为 OfficeThemes.css。
    

### <a name="referencing-officethemescss-in-your-add-ins-html-pages"></a>在加载项的 HTML 页中引用 OfficeThemes.css

若要在你的外接程序项目中使用 OfficeThemes.css，则在 Web 页（如 .html、.aspx 或 .php 文件）的 `<head>` 标记内添加引用 OfficeThemes.css 文件的 `<link>` 标记，这些 Web 页使用以下格式实现外接程序的 UI：


```HTML
<link href="<local_path_to_OfficeThemes.css> " rel="stylesheet" type="text/css" />
```

要在 Visual Studio 中执行此操作，请按照下面的步骤操作。


### <a name="to-reference-officethemescss-in-your-add-in-for-powerpoint"></a>在 PowerPoint 加载项中引用 OfficeThemes.css


1. 在 Visual Studio 2015 中，打开或创建新的“**Office 外接程序**”项目。
    
2. 在实现外接程序 UI 的 HTML 页（如默认模板中的 Home.html）中，在引用 OfficeThemes.css 文件的 `<head>` 标记内添加以下 `<link>` 标记：
    
```HTML
  <link href="../../Content/OfficeThemes.css" rel="stylesheet" type="text/css" />
```

如果你正在使用非 Visual Studio 的工具创建外接程序，请添加相同格式的 `<link>` 标记，这一标记指定了将使用外接程序进行部署的 OfficeThemes.css 副本的相对路径。


### <a name="using-officethemescss-document-theme-classes-in-your-content-add-ins-html-page"></a>在内容加载项的 HTML 页中使用 OfficeThemes.css 文档主题类

以下演示了使用 OfficeTheme.css 文档主题类的内容 外接程序 中的 HTML 简单示例。有关与文档主题中 12 种颜色和 2 种字体对应的 OfficeThemes.css 类的详细信息，请参阅 [适用于内容加载项的主题类](#theme-classes-for-content-add-ins)。


```HTML
<body> <div id="themeSample" class="office-docTheme-primary-fontColor "> <h1 class="office-headerFont-latin">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent1-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent2-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent3-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent4-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent5-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent6-bgColor">Hello world!</h1> <p class="office-bodyFont-latin office-docTheme-secondary-fontColor">Hello world!</p> </div> </body>
```

运行时，当在使用默认“**Office 主题**”的演示文稿中插入内容外接程序时，呈现如下：


![运行 Office 主题的内容应用程序](../../images/off15app_ContentApp_OfficeTheme.png)

如果你将演示文稿更改为使用其他主题或自定义该演示文稿的主题，由 OfficeThemes.css 类指定的字体和颜色将动态更新为与演示文稿主题的字体和颜色对应。使用与上述相同的 HTML 示例时，如果外接程序插入的演示文稿使用**Facet**主题，外接程序将呈现如下：


![运行 Facet 主题的内容应用程序](../../images/off15app_ContentApp_FacetTheme.png)


### <a name="using-officethemescss-office-ui-theme-classes-in-your-task-pane-add-ins-html-page"></a>在任务窗格加载项的 HTML 页中使用 OfficeThemes.css Office UI 主题类

除文档主题之外，用户还可为所有 Office 应用程序自定义 Office 用户界面的颜色主题，通过使用“**文件**” > “**帐户**” > “**Office 主题**”下拉框进行自定义。

以下演示了 HTML 的简单示例，该示例在任务窗格 外接程序 中使用 OfficeTheme.css 类指定字体颜色和背景色。有关与 Office UI 主题字体和颜色对应的 OfficeThemes.css 类的详细信息，请参阅 [适用于任务窗格加载项的主题类](#theme-classes-for-task-pane-add-ins)。


```HTML
<body> <div id="content-header" class="office-officeTheme-primary-fontColor office-officeTheme-primary-bgColor"> <div class="padding"> <h1>Welcome</h1> </div> </div> <div id="content-main" class="office-officeTheme-secondary-fontColor office-officeTheme-secondary-bgColor"> <div class="padding"> <p>Add home screen content here.</p> <p>For example:</p> <button id="get-data-from-selection">Get data from selection</button> <p> <a target="_blank" class="office-a" href="https://go.microsoft.com/fwlink/?LinkId=276812"> Find more samples online... </a> </p> </div> </div> </body> 
```

当在 PowerPoint 中运行时，且“**文件**” > “**帐户** > “**Office 主题**”设置为“**白色**”时，任务窗格外接程序呈现如下：


![使用 Office 白色主题的任务窗格](../../images/off15app_TaskPaneThemeWhite.png)

如果你将 **Office 主题**更改为**深灰色**，由 OfficeThemes.css 类指定的字体和颜色将动态更新，呈现如下：


![使用 Office 深灰色主题的任务窗格](../../images/off15app_TaskPaneThemeDarkGray.png)


## <a name="officethemecss-classes"></a>OfficeTheme.css 类


OfficeThemes.css 文件包括两组类，您可用于 PowerPoint 内容和任务窗格加载项。


### <a name="theme-classes-for-content-add-ins"></a>适用于内容加载项的主题类


OfficeThemes.css 文件提供与文档主题中的 2 种字体和 12 种颜色对应的类。这些类很适合用于 PowerPoint 内容加载项，以便您的加载项字体和颜色与它要插入的演示文稿相协调。


**适用于内容外接程序的主题字体**


|**类**|**说明**|
|:-----|:-----|
| `office-bodyFont-eastAsian`|正文字体的东亚名称。|
| `office-bodyFont-latin`|正文字体的拉丁名称。默认为"Calabri"|
| `office-bodyFont-script`|正文字体的脚本名称。|
| `office-bodyFont-localized`|正文字体的本地化名称。根据当前在 Office 中使用的区域性，指定默认字体名称。|
| `office-headerFont-eastAsian`|标题字体的东亚名称。|
| `office-headerFont-latin`|标题字体的拉丁名称。默认为"Calabri Light"|
| `office-headerFont-script`|标题字体的脚本名称。|
| `office-headerFont-localized`|标题字体的本地化名称。根据当前在 Office 中使用的区域性，指定默认字体名称。|

**适用于内容外接程序的主题颜色**


|**类**|**说明**|
|:-----|:-----|
| `office-docTheme-primary-fontColor`|首选字体颜色。默认为 #000000|
| `office-docTheme-primary-bgColor`|首选字体背景色。默认为 #FFFFFF|
| `office-docTheme-secondary-fontColor`|辅助字体颜色。默认为 #000000|
| `office-docTheme-secondary-bgColor`|辅助字体背景色。默认为 #FFFFFF|
| `office-contentAccent1-color`|字体个性色 1。默认为 #5B9BD5|
| `office-contentAccent2-color`|字体个性色 2。默认为 #ED7D31|
| `office-contentAccent3-color`|字体个性色 3。默认为 #A5A5A5|
| `office-contentAccent4-color`|字体个性色 4。默认为 #FFC000|
| `office-contentAccent5-color`|字体个性色 5。默认为 #4472C4|
| `office-contentAccent6-color`|字体个性色 6。默认为 #70AD47|
| `office-contentAccent1-bgColor`|背景个性色 1。默认为 #5B9BD5|
| `office-contentAccent2-bgColor`|背景个性色 2。默认为 #ED7D31|
| `office-contentAccent3-bgColor`|背景个性色 3。默认为 #A5A5A5|
| `office-contentAccent4-bgColor`|背景个性色 4。默认为 #FFC000|
| `office-contentAccent5-bgColor`|背景个性色 5。默认为 #4472C4|
| `office-contentAccent6-bgColor`|背景个性色 6。默认为 #70AD47|
| `office-contentAccent1-borderColor`|边框个性色 1。默认为 #5B9BD5|
| `office-contentAccent2-borderColor`|边框个性色 2。默认为 #ED7D31|
| `office-contentAccent3-borderColor`|边框个性色 3。默认为 #A5A5A5|
| `office-contentAccent4-borderColor`|边框强调文字颜色 4。默认为 #FFC000|
| `office-contentAccent5-borderColor`|边框个性色 5。默认为 #4472C4|
| `office-contentAccent6-borderColor`|边框个性色 6。默认为 #70AD47|
| `office-a`|超链接颜色。默认为 #0563C1|
| `office-a:visited`|已访问的超链接颜色。默认为 #954F72|
以下屏幕截图显示，在使用默认 Office 主题时，分配给 外接程序 文本的所有主题颜色类（两种超链接颜色除外）的示例。


![默认 Office 主题颜色示例](../../images/off15app_DefaultOfficeThemeColors.png)


### <a name="theme-classes-for-task-pane-add-ins"></a>适用于任务窗格加载项的主题类


OfficeThemes.css 文件提供的类与分配给 Office 应用程序 UI 主题所使用的字体和背景的 4 种颜色对应。这些类很适合用于 PowerPoint 相关的任务加载项，以便您的加载项颜色与其他 Office 内置的任务窗格协调。


**适用于任务窗格外接程序的主题字体和背景色**


|**类**|**说明**|
|:-----|:-----|
| `office-officeTheme-primary-fontColor`|首选字体颜色。默认为：#B83B1D|
| `office-officeTheme-primary-bgColor`|首选背景色。默认为 #DEDEDE|
| `office-officeTheme-secondary-fontColor`|辅助字体颜色。默认为 #262626|
| `office-officeTheme-secondary-bgColor`|辅助背景色。默认为 #FFFFFF|

## <a name="additional-resources"></a>其他资源

- [创建 PowerPoint 相关的内容和任务窗格外接程序](../powerpoint/powerpoint-add-ins.md)
