---
title: 清单文件中的 Version 元素
description: Version 元素指定Office外接程序版本。
ms.date: 02/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 34cefa22123ed4ee723d51a669e01e042efc2934
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152587"
---
# <a name="version-element"></a>Version 元素

指定 Office 外接程序的版本。 版本号可以是 1、2、3 或 4 (，即 n、n.n、n.n 或 n.n.n) 。

**加载项类型：** 内容、任务窗格和邮件

## <a name="syntax"></a>语法

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>包含于

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注解

版本号的每个部分最多可包含 5 个数字。
