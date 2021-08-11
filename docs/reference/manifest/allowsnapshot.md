---
title: 清单文件中的 AllowSnapshot 元素
description: 指定是否将内容外接程序的快照图像与主机文档一起保存。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 1462b60dffda7e3bb611225f015b5a1c9f0b5e78271580383961cc118af60587
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095051"
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
 > **AllowSnapshot** 在默认情况下为 `true`。 这样，对于在不支持 Office 加载项的 Office 应用程序版本中打开文档的用户提供的加载项图像，或者如果应用程序无法连接到托管加载项的服务器，则提供加载项的静态图像。 但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。
