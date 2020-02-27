---
title: 清单文件中的运行时（预览）
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 17e53b53d55ea9547cdfc5c4f89f8f4c3a7ab75e
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283870"
---
# <a name="runtimes-element-preview"></a>运行时元素（预览）

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

指定外接程序的运行时，并启用自定义函数、功能区按钮和任务窗格，以使用相同的 JavaScript 运行时。 清单文件中`<Host>`的元素的子元素。 有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。

**外接程序类型：** 任务窗格

> [!IMPORTANT]
> 共享运行时当前处于预览阶段，仅适用于 Windows 上的 Excel。 若要尝试预览功能，你需要加入[Office 预览体验成员](https://insider.office.com/)。

## <a name="syntax"></a>语法

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>包含于 
[Host](./host.md)

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **运行时**     | 是 |  外接程序的运行时。

## <a name="see-also"></a>另请参阅

- [运行时](runtime.md)
