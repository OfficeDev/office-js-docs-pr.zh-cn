---
title: Office 外接程序主机和平台可用性
description: Excel、OneNote、Outlook、PowerPoint、Project 和 Word 支持的要求集。
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 8c3c187d8f9b70f40a35e3773a2267dc76decbd0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611980"
---
# <a name="office-add-in-host-and-platform-availability"></a>Office 外接程序主机和平台可用性

若要按预期运行，Office 加载项可能会依赖特定的 Office 主机、要求集、API 成员或 API 版本。下表列出了每个 Office 应用目前支持的可用平台、扩展点、API 要求集和通用 API。

> [!NOTE]
> 通过 MSI 安装的初始 Office 2016 版本只包含 ExcelApi 1.1、WordApi 1.1 和通用 API 要求集。 有关各种 Office 版本更新历史记录的更多信息，请查看[另请参阅](#see-also)部分。

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">平台</th>
    <th style="width:10%">扩展点</th>
    <th style="width:20%">API 要求集</th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Office 网页版</td>
    <td> - 任务窗格<br>
        - 内容<br>
        - 自定义函数<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - 任务窗格<br>
        - 内容<br>
        - 自定义函数<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>（一次性购买）</td>
    <td>- 任务窗格<br>
        - 内容<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>（一次性购买）</td>
    <td>- 任务窗格<br>
        - 内容</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>（一次性购买）</td>
    <td>
        - 任务窗格<br>
        - 内容</td>
    <td>  - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>iPad 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td>- 任务窗格<br>
        - 内容</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td>- 任务窗格<br>
        - 内容<br>
        - 自定义函数<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 版 Office 2019<br>（一次性购买）</td>
    <td>- 任务窗格<br>
        - 内容<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">加载项命令</a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 版 Office 2016<br>（一次性购买）</td>
    <td>- 任务窗格<br>
        - 内容</td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
</table>

*&ast; - 已添加发布后更新。*

## <a name="custom-functions-excel-only"></a>自定义函数（仅 Excel）

<table style="width:80%">
  <tr>
    <th style="width:10%">平台</th>
    <th style="width:10%">扩展点</th>
    <th style="width:20%">API 要求集</th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Office 网页版</td>
    <td>
        - 自定义函数</td>
    <td>
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
  </tr>
  <tr>
    <td>Windows 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td>
        - 自定义函数</td>
    <td>
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
  </tr>
  <tr>
    <td>Office for Mac<br>（连接到 Office 365）</td>
    <td>
        - 自定义函数</td>
    <td>
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Office 网页版<br>（新式）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Office 网页版<br>（经典）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Windows 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>（一次性购买）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>（一次性购买）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">模块</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>（一次性购买）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*<br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>iOS 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Mac 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Mac 版 Office 2019<br>（一次性购买）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Mac 版 Office 2016<br>（一次性购买）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">邮件撰写</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">约会参与者（阅读）</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">约会参与者（撰写）</a><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>不可用</td>
  </tr>
  <tr>
    <td>Android 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">邮件阅读</a><br>
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">约会组织者（撰写）：联机会议</a> （预览）<br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>不可用</td>
  </tr>
</table>

*&ast; - 已添加发布后更新。*

> [!IMPORTANT]
> 要求集的客户端支持可能受到 Exchange 服务器支持的限制。 有关 Exchange 服务器和 Outlook 客户端支持的要求集范围的详细信息，请参阅 [Outlook JavaScript API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Office 网页版</td>
    <td> - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>Windows 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>（一次性购买）</td>
    <td> - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>（一次性购买）</td>
    <td> - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>（一次性购买）</td>
    <td> - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>iPad 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Mac 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Mac 版 Office 2019<br>（一次性购买）</td>
    <td> - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Mac 版 Office 2016<br>（一次性购买）</td>
    <td> - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
</table>

*&ast; - 已添加发布后更新。*

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Office 网页版</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>（一次性购买）</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>（一次性购买）</td>
    <td> - 内容<br>
         - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>（一次性购买）</td>
    <td> - 内容<br>
         - 任务窗格<br>
    </td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>iPad 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - 内容<br>
         - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 版 Office<br>（已连接到 Office 365 订阅）</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 版 Office 2019<br>（一次性购买）</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Mac 版 Office 2016<br>（一次性购买）</td>
    <td> - 内容<br>
         - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
</table>

*&ast; - 已添加发布后更新。*

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Office 网页版</td>
    <td> - 内容<br>
         - 任务窗格<br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">外接程序命令</a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></td>
    <td> - DocumentEvents<br>
         - HtmlCoercion<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="project"></a>项目

<table style="width:80%">
  <tr>
    <th>平台</th>
    <th>扩展点</th>
    <th>API 要求集</th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>通用 API</b></a></th>
  </tr>
  <tr>
    <td>Windows 版 Office 2019<br>（一次性购买）</td>
    <td> - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2016<br>（一次性购买）</td>
    <td> - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Windows 版 Office 2013<br>（一次性购买）</td>
    <td> - 任务窗格</td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - Selection<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](office-add-ins.md)
- [Office 版本和要求集](../develop/office-versions-and-requirement-sets.md)
- [通用 API 要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [加载项命令要求集](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [API 参考文档](../reference/javascript-api-for-office.md)
- [Office 365 ProPlus 的更新历史记录](/officeupdates/update-history-office365-proplus-by-date)
- [Office 2016 和 2019 更新历史记录（即点即用）](/officeupdates/update-history-office-2019)
- [Office 2013 更新历史记录（即点即用）](/officeupdates/update-history-office-2013)
- [Office 2010、2013 和 2016 更新历史记录 (MSI)](/officeupdates/office-updates-msi)
- [Outlook 2010、2013 和 2016 更新历史记录 (MSI)](/officeupdates/outlook-updates-msi)
- [Office for Mac 更新历史记录](/officeupdates/update-history-office-for-mac)
- [构建 Office 加载项](../overview/office-add-ins-fundamentals.md)