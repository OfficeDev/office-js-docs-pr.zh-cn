---
title: 清单文件中的运行时
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554004"
---
# <a name="runtimes-element"></a>运行时元素

此功能处于预览阶段。 指定外接程序的运行时，并允许自定义函数和任务窗格共享全局数据，并使函数相互调用。 应遵循清单`<Host>`文件中的元素。

**外接程序类型：** 任务窗格

## <a name="syntax"></a>语法

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **运行时**     | 是 |  外接程序的运行时通常与 Excel 自定义函数一起使用。

## <a name="see-also"></a>另请参阅

- [运行时](runtime.md)
