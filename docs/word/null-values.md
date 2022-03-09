---
title: Word 加载项中的 Null 值
description: 了解如何在 Word 加载项中处理空值。
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e21677dafcaaaa7e9e9164ef18c82f49820298d6
ms.sourcegitcommit: 9d930b4c77c342246607aef30479e31fdbdd47f0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63353855"
---
# <a name="null-values-in-word-add-ins"></a>Word 加载项中的 Null 值

`null` 在 Word JavaScript API 中具有特殊含义。 它用于表示默认值或无格式。

## <a name="null-property-values-in-the-response"></a>响应中的 null 属性值

当指定 [区域存在不同](/javascript/api/word/word.font#word-word-font-color-member) 值 `null` 时，颜色等格式属性将包含响应中的 [值](/javascript/api/word/word.range)。 例如，如果你检索某个区域并加载其 `range.font.color` 属性：

- 如果区域内的所有文本都具有相同的字体颜色， `range.font.color` 则指定该颜色。
- 如果该区域内存在多种字体颜色，则 `range.font.color` 为 `null`。
