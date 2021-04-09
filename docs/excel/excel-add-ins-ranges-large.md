---
title: 使用 Excel JavaScript API 读取或写入较大区域
description: 了解如何使用 Excel JavaScript API 读取或写入较大区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b7a1e54d6b516889884f777bd256df8fb663c794
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652797"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 读取或写入较大区域

本文介绍如何使用 Excel JavaScript API 处理对较大范围的读取和写入。

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a>对较大区域运行单独的读取或写入操作

如果某个区域包含大量单元格、值、数字格式或公式，则可能无法对区域运行 API 操作。 API 将始终尽量尝试在区域内运行所请求的操作（即检索或写入指定的数据），但尝试对较大区域执行读取或写入操作可能会因资源利用率过高而导致 API 错误。 为避免此类错误，建议为较大区域的较小子集运行单独的读取或写入操作，而不是尝试在较大区域内运行单个读取或写入操作。

有关系统限制的详细信息，请参阅 Office 加载项的资源限制和性能优化的 ["Excel 加载项"部分](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)。

### <a name="conditional-formatting-of-ranges"></a>范围的条件格式

范围可以根据条件将格式应用于个别单元格。 有关此操作的详细信息，请参阅[将条件格式应用于 Excel 范围](excel-add-ins-conditional-formatting.md)。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [使用 Excel JavaScript API 读取或写入无限区域](excel-add-ins-ranges-unbounded.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
