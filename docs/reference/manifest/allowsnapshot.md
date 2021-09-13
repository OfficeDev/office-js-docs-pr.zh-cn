---
title: 清单文件中的 AllowSnapshot 元素
description: 指定是否将内容外接程序的快照图像与主机文档一起保存。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 723817557020f4ec3dbe5b3135877fe49bf67acb
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151932"
---
# <a name="allowsnapshot-element"></a>AllowSnapshot 元素

指定是否将内容外接程序的快照图像与主机文档一起保存。

**外接程序类型：** 内容

## <a name="syntax"></a>语法

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注释

 > [!IMPORTANT]
 > **AllowSnapshot** 在默认情况下为 `true`。 这样，如果用户在不支持 Office 外接程序的 Office 应用程序版本中打开文档，或者该应用程序无法连接到托管该外接程序的服务器，则提供外接程序的静态图像，则用户可以看到该外接程序的图像。 但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。
