# <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

要求集是指各组已命名的 API 成员。 Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。 有关详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。

需要了解加载项在哪些位置受 Office 主机支持？ 请参阅 [Office 加载项主机和平台可用性](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)。

正在寻找*主机专用* API 要求集吗？ 请参阅下列 API 要求集：
 
- [Excel JavaScript API 要求集](excel-api-requirement-sets.md) (ExcelApi)
- [Word JavaScript API 要求集](word-api-requirement-sets.md) (WordApi)
- [OneNote JavaScript API 要求集](onenote-api-requirement-sets.md) (OneNoteApi)
- [了解 Outlook API 要求集](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> 我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。 作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。

## <a name="common-api-requirement-sets"></a>通用 API 要求集

下表列出了通用 API 要求集、每个集内的方法，以及支持相应要求集的 Office 主机应用程序。这些 API 要求集都是第 1.1 版。

|**要求集**|**Office 主机**|**要求集内的方法**|
|:-----|:-----|:-----|
| ActiveView | PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac|Document.getActiveViewAsync|
| AddInCommands | 请参阅[加载项命令要求集](add-in-commands-requirement-sets.md)。 | |
| BindingEvents  | Access Web 应用<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持使用 Document.getFileAsync 方法时输出作为字节数组 (Office.FileType.Compressed) 的 Office Open XML (OOXML) 格式<br>。|
| CustomXmlParts    | Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogApi | 请参阅 [Dialog API 要求集](dialog-api-requirement-sets.md)。 | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |
| DocumentEvents    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| 文件  | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | OneNote Online<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为 HTML (Office.CoercionType.Html)<br>。|
| IdentityAPI | 请参阅 [Identity API 要求集](identity-api-requirement-sets.md)。 | Auth.getAccessTokenAsync |
| ImageCoercion | Excel<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持在使用 Document.setSelectedDataAsync 方法写入数据时转换为图像 (Office.CoercionType.Image)。|
| 邮箱   |Outlook for Windows<br>Outlook for web<br>Outlook for Android<br>Outlook for Mac<br>Outlook Web App |请参阅[了解 Outlook API 要求集](outlook-api-requirement-sets.md)。|
| MatrixBindings    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word<br>Word Online<br>Word for iPad<br>Word for Mac|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为“矩阵”（数组的数组）数据结构 (Office.CoercionType.Matrix)。|
| OoxmlCoercion | Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为 Open Office XML (OOXML) 格式 (Office.CoercionType.Ooxml)。|
| PartialTableBindings  | Access Web 应用||
| PdfFile   | Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持使用 Document.getFileAsync 方法时输出 PDF 格式 (Office.FileType.Pdf)<br>。|
| Selection | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Project<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Settings  | Access Web 应用<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | Access Web 应用<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | Access Web 应用<br>Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为“表格”数据结构 (Office.CoercionType.Table)。|
| TextBindings  | Excel<br>Excel Online<br>Excel for iPad<br>Excel for Mac<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | Excel<br>Excel Online<br>Excel for iPad<br>OneNote Online<br>PowerPoint<br>PowerPoint Online<br>PowerPoint for iPad<br>PowerPoint for Mac<br>Project<br>Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为文本格式 (Office.CoercionType.Text)。|
| TextFile  | Word 2013 及更高版本<br>Word 2016 for Mac 及更高版本<br>Word Online<br>Word for iPad|支持在使用 Document.getFileAsync 方法时输出文本格式 (Office.FileType.Text)。|

## <a name="methods-that-arent-part-of-a-requirement-set"></a>不作为要求集一部分的方法

适用于 Office 的 JavaScript API 中的以下方法不是要求集的一部分。 如果加载项需要这些方法的任意一个，请使用加载项清单中的 **Methods** 和 **Method** 元素以声明需要这些方法，或使用 `if` 语句执行运行时检查。 有关详细信息，请参阅[指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)。

|**方法名称**|**Office 主机支持**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access web app、Excel、Excel Online、Excel for iPad 和 Excel for Mac|
|Document.getFilePropertiesAsync|Excel、Excel Online、Excel for iPad、Excel for Mac、PowerPoint、PowerPoint Online、PowerPoint for iPad、PowerPoint for Mac、Word、Word Online、Word for iPad 和 Word for Mac|
|Document.getProjectFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.goToByIdAsync|Excel、Excel Online、Excel for iPad、Excel for Mac、PowerPoint、PowerPoint Online、PowerPoint for iPad、PowerPoint for Mac、Word、Word Online、Word for iPad 和 Word for Mac|
|Settings.addHandlerAsync|Access Web App 和 Excel Online|
|Settings.refreshAsync|Access web app、Excel、Excel Online、PowerPoint、PowerPoint Online、Word 和 Word Online|
|Settings.removeHandlerAsync|Access Web App 和 Excel Online|
|TableBinding.clearFormatsAsync|Excel、Excel Online、Excel for iPad 和 Excel for Mac|
|TableBinding.setFormatsAsync|Excel、Excel Online、Excel for iPad 和 Excel for Mac|
|TableBinding.setTableOptionsAsync|Excel、Excel Online、Excel for iPad 和 Excel for Mac|

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office 外接程序 XML 清单](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
