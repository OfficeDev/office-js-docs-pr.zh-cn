---
title: Office 加载项主机和平台可用性
description: Excel、Word、Outlook、PowerPoint 和 OneNote 支持的要求集。
ms.date: 08/30/2018
ms.openlocfilehash: 06fb073693bd8adca7d196f4361699ac3f54cee1
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797299"
---
# <a name="office-add-in-host-and-platform-availability"></a>Office 加载项主机和平台可用性

若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。 下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API 要求集。

如果表格单元格内有星号 (*)，表示我们正在完善它。 有关 Project 或 Access 要求集，请参阅 [Office 通用要求集](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)。  

> [!NOTE]
> 通过 MSI 安装的 Office 2016 的生成号为 16.0.4266.1001。此版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">平台</th>
    <th style="width:10%">扩展点</th>
    <th style="width:20%">API 要求集</th>
    <th style="width:40%"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - 任务窗格<br>
        - 内容<br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </td>
    <td>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - 文件<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - 选择<br>
        - 设置<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td>
        - 任务窗格<br>
        - 内容</td>
    <td>  - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - 文件<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - 选择<br>
        - 设置<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td>- 任务窗格<br>
        - 内容<br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td>- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - 文件<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - 选择<br>
        - 设置<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td>- 任务窗格<br>
        - 内容</td>
    <td>- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - 文件<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - 选择<br>
        - 设置<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td>- 任务窗格<br>
        - 内容<br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td>- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - 文件<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - 选择<br>
        - 设置<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - 邮件阅读<br>
      - 邮件撰写<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>不适用</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td> - 邮件阅读<br>
      - 邮件撰写<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></td>
    <td>不适用</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td> - 邮件阅读<br>
      - 邮件撰写<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a><br>
      - 模块</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>不适用</td>
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td> - 邮件阅读<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>不适用</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td> - 邮件阅读<br>
      - 邮件撰写<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>不适用</td>
  </tr>
  <tr>
    <td>Office for Android</td>
    <td> - 邮件阅读<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>不适用</td>
  </tr>
</table>

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - 任务窗格<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - 文件<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td> - 任务窗格</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - 文件<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td> - 任务窗格<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - 文件<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td> - 任务窗格</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - 文件<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td> - 任务窗格<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - 文件<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - 文件<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 for Windows</td>
    <td> - 内容<br>
         - 任务窗格<br>
    </td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - 文件<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 for Windows</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - 文件<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office for iOS</td>
    <td> - 内容<br>
         - 任务窗格</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
     <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - 文件<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 for Mac</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - 文件<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - 选择<br>
         - 设置<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - DocumentEvents<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - 设置<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](office-add-ins.md)
- [通用 API 要求集](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [加载项命令要求集](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [适用于 Office 的 JavaScript API 参考](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
