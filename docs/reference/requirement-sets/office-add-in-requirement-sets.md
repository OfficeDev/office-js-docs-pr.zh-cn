---
title: Office 通用 API 要求集
description: 了解有关 Office 通用 API 要求集的详细信息。
ms.date: 09/17/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d5fd33a2c44cb85e8279a970d4d7443783f049ff
ms.sourcegitcommit: 2479812e677d1a7337765fe8f1c8345061d4091a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/19/2020
ms.locfileid: "48135219"
---
# <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

> [!TIP]
> 是否要查找 *特定于应用程序* 的 API 要求集？ 请参阅下列 API 要求集：
>
> - [Excel JavaScript API 要求集](excel-api-requirement-sets.md) (ExcelApi)
> - [Word JavaScript API 要求集](word-api-requirement-sets.md) (WordApi)
> - [OneNote JavaScript API 要求集](onenote-api-requirement-sets.md) (OneNoteApi)
> - [PowerPoint JavaScript API 要求集](powerpoint-api-requirement-sets.md) (PowerPointApi)
> - [了解 Outlook API 要求集](outlook-api-requirement-sets.md) (MailBox)

> [!IMPORTANT]
> 我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。 作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。

## <a name="common-api-requirement-sets"></a>通用 API 要求集

以下各节列出了常见的 API 要求集、每个集合中的方法以及支持该要求集的 Office 客户端应用程序。 除非另行指定，否则这些 API 要求集都是第 1.1 版。

> [!TIP]
> 需要有关 Office 应用程序和版本支持加载项和要求集的信息？ [有关 Office 加载项，请参阅 office 客户端应用程序和平台可用性](../../overview/office-add-in-availability.md)。

### <a name="activeview"></a>ActiveView

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 PowerPoint<br>PowerPoint 网页版<br>iPad 版 PowerPoint<br>Mac 版 PowerPoint|Document.getActiveViewAsync|

---

### <a name="addincommands"></a>AddInCommands

请参阅[加载项命令要求集](add-in-commands-requirement-sets.md)。

---

### <a name="bindingevents"></a>BindingEvents

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Access Web 应用<br>Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|Binding.addHandlerAsync<br>Binding.removeHandlerAsync|

---

### <a name="compressedfile"></a>CompressedFile

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Excel 2016 及更高版本的 Windows<br>Excel 网页版<br>Excel 2016 及更高版本 Mac<br>Windows 版 PowerPoint<br>PowerPoint 网页版<br>iPad 版 PowerPoint<br>Mac 版 PowerPoint<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|支持使用 Document.getFileAsync 方法时输出作为字节数组 (Office.FileType.Compressed) 的 Office Open XML (OOXML) 格式<br>。|

---

### <a name="customxmlparts"></a>CustomXmlParts

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getTextAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setTextAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|

---

### <a name="dialogapi"></a>DialogApi

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| 请参阅 [Dialog API 要求集](dialog-api-requirement-sets.md)。 | UI.messageParent<br>UI.displayDialogAsync<br>UI.closeContainer<br>UI.Dialog |

---

### <a name="documentevents"></a>DocumentEvents

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>OneNote 网页版<br>Windows 版 PowerPoint<br>PowerPoint 网页版<br>iPad 版 PowerPoint<br>Mac 版 PowerPoint<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|Document.addHandlerAsync<br>Document.removeHandlerAsync|

---

### <a name="file"></a>文件

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>Windows 版 PowerPoint<br>PowerPoint 网页版<br>iPad 版 PowerPoint<br>Mac 版 PowerPoint<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|

---

### <a name="htmlcoercion"></a>HtmlCoercion

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| OneNote 网页版<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为 HTML (Office.CoercionType.Html)。|

---

### <a name="identityapi"></a>IdentityAPI

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| 请参阅 [Identity API 要求集](identity-api-requirement-sets.md)。 | Auth.getAccessToken |

---

### <a name="imagecoercion"></a>ImageCoercion

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| 请参阅[图像强制要求集](image-coercion-requirement-sets.md)。 | Document.setSelectedDataAsync 方法|

---

### <a name="mailbox"></a>Mailbox

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
|Windows 版 Outlook<br>Outlook 网页版<br>Android 版 Outlook<br>Mac 版 Outlook<br>iOS 版 Outlook|请参阅[了解 Outlook API 要求集](outlook-api-requirement-sets.md)。|

---

### <a name="matrixbindings"></a>MatrixBindings

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>Windows 版 Word<br>Word 网页版<br>iPad 版 Word<br>Mac 版 Word|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="matrixcoercion"></a>MatrixCoercion

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为“矩阵”（数组的数组）数据结构 (Office.CoercionType.Matrix)。|

---

### <a name="ooxmlcoercion"></a>OoxmlCoercion

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为 Open Office XML (OOXML) 格式 (Office.CoercionType.Ooxml)。|

---

### <a name="openbrowserwindowapi"></a>OpenBrowserWindowApi

|**Office 主机**|**要求集内的方法**|
|:-----|:-----|
| 请参阅 [打开浏览器窗口 API 要求集](open-browser-window-api-requirement-sets.md)。 | OpenBrowserWindow 的用户 |

---

### <a name="partialtablebindings"></a>PartialTableBindings

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Access Web 应用||

---

### <a name="pdffile"></a>PdfFile

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Excel<br>Excel 网页版<br>Mac 版 Excel<br>Windows 版 PowerPoint<br>PowerPoint 网页版<br>iPad 版 PowerPoint<br>Mac 版 PowerPoint<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|支持使用 Document.getFileAsync 方法时输出 PDF 格式 (Office.FileType.Pdf)<br>。|

---

### <a name="ribbonapi"></a>RibbonApi

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| 请参阅 [功能区 API 要求集](ribbon-api-requirement-sets.md)。 | RequestUpdate |

---

### <a name="selection"></a>Selection

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>Windows 版 PowerPoint<br>PowerPoint 网页版<br>iPad 版 PowerPoint<br>Mac 版 PowerPoint<br>Windows 版 Project<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|

---

### <a name="settings"></a>Settings

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Access Web 应用<br>Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>OneNote 网页版<br>Windows 版 PowerPoint<br>PowerPoint 网页版<br>iPad 版 PowerPoint<br>Mac 版 PowerPoint<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|

---

### <a name="sharedruntime"></a>SharedRuntime

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| 请参阅 [共享运行时要求集](shared-runtime-requirement-sets.md)。 | GetStartupBehavior<br>.Addin： hide<br>OnVisibilityModeChanged<br>SetStartupBehavior<br>ShowAsTaskpane<br> |

---

### <a name="tablebindings"></a>TableBindings

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Access Web 应用<br>Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.addColumnsAsync<br>Binding.addRowsAsync<br>Binding.deleteAllDataValuesAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="tablecoercion"></a>TableCoercion

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Access Web 应用<br>Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为“表格”数据结构 (Office.CoercionType.Table)。|

---

### <a name="textbindings"></a>TextBindings

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>Mac 版 Excel<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsync<br>Binding.getDataAsync<br>Binding.setDataAsync|

---

### <a name="textcoercion"></a>TextCoercion

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Excel<br>Excel 网页版<br>iPad 版 Excel<br>OneNote 网页版<br>Windows 版 PowerPoint<br>PowerPoint 网页版<br>iPad 版 PowerPoint<br>Mac 版 PowerPoint<br>Windows 版 Project<br>Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|支持在使用 Document.getSelectedDataAsync、Document.setSelectedDataAsync、Binding.getDataAsync 或 Binding.setDataAsync 方法读取和写入数据时强制转换为文本格式 (Office.CoercionType.Text)。|

---

### <a name="textfile"></a>TextFile

|**Office 应用程序**|**要求集内的方法**|
|:-----|:-----|
| Windows 版 Word 2013 及更高版本<br>Mac 版 Word 2016 及更高版本<br>Word 网页版<br>iPad 版 Word|支持在使用 Document.getFileAsync 方法时输出文本格式 (Office.FileType.Text)。|

---

## <a name="methods-that-arent-part-of-a-requirement-set"></a>不作为要求集一部分的方法

Office JavaScript API 中的以下方法不是要求集的一部分。 如果加载项需要这些方法的任意一个，请使用加载项清单中的 **Methods** 和 **Method** 元素以声明需要这些方法，或使用 `if` 语句执行运行时检查。 有关详细信息，请参阅 [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)。

|**方法名称**|**Office 应用程序支持**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Access Web 应用、Windows 版 Excel、Excel 网页版、iPad 版 Excel 和 Mac 版 Excel|
|Document.getFilePropertiesAsync|Windows 版 Excel、Excel 网页版、iPad 版 Excel、Mac 版 Excel、Windows 版 PowerPoint、PowerPoint 网页版、iPad 版 PowerPoint、Mac 版 PowerPoint、Windows 版 Word、Word 网页版、iPad 版 Word 和 Mac 版 Word|
|Document.getProjectFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getResourceFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getSelectedViewAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getTaskAsync|Project Standard 2013 和 Project Professional 2013|
|Document.getTaskFieldAsync|Project Standard 2013 和 Project Professional 2013|
|Document.goToByIdAsync|Windows 版 Excel、Excel 网页版、iPad 版 Excel、Mac 版 Excel、Windows 版 PowerPoint、PowerPoint 网页版、iPad 版 PowerPoint、Mac 版 PowerPoint、Windows 版 Word、Word 网页版、iPad 版 Word 和 Mac 版 Word|
|Settings.addHandlerAsync|Access Web 应用和 Excel 网页版|
|Settings.refreshAsync|Access Web 应用、Windows 版 Excel、Excel 网页版、Windows 版 PowerPoint、PowerPoint 网页版、Word 和 Word 网页版|
|Settings.removeHandlerAsync|Access Web 应用和 Excel 网页版|
|TableBinding.clearFormatsAsync|Windows 版 Excel、Excel 网页版、iPad 版 Excel 和 Mac 版 Excel|
|TableBinding.setFormatsAsync|Windows 版 Excel、Excel 网页版、iPad 版 Excel 和 Mac 版 Excel|
|TableBinding.setTableOptionsAsync|Windows 版 Excel、Excel 网页版、iPad 版 Excel 和 Mac 版 Excel|

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 应用程序和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
