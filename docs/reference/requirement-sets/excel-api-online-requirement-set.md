---
title: Excel JavaScript API 仅联机要求集
description: 有关 ExcelApiOnline 要求集的详细信息
ms.date: 11/19/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e583c9832f04e17dc1c82d38d056fe2749888a77
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757490"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API 仅联机要求集

`ExcelApiOnline`要求集是一个特殊要求集，其中包含仅适用于 web 上的 Excel 的功能。 此要求集中的 Api 被认为是针对 web 主机上的 Excel 的生产 Api （不受未记录的行为或结构更改）。 `ExcelApiOnline`被视为针对其他平台（Windows、Mac、iOS）的 "预览" Api，这些平台可能不支持这些平台。

当在所有平台`ExcelApiOnline`上支持要求集中的 api 时，它们将添加到下一个发布的要求集`ExcelApi 1.[NEXT]`（）。 一旦新要求是公共的，将从这些 Api 中`ExcelApiOnline`删除。 可将此视为将 API 从预览迁移到发布的类似升级过程。

> [!IMPORTANT]
> `ExcelApiOnline`是最新编号的要求集的超集。

> [!IMPORTANT]
> `ExcelApiOnline 1.1`是仅联机 Api 的唯一版本。 这是因为 web 上的 Excel 将始终有一个版本可供最新版本的用户使用。

## <a name="recommended-usage"></a>建议使用

由于`ExcelApiOnline` web 上的 Excel 仅支持 api，因此，您的外接程序应检查是否支持要求集，然后再调用这些 api。 这样可以避免在不同的平台上调用仅联机 API。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

一旦 API 位于跨平台要求集，就应删除或编辑该`isSetSupported`检查。 这将在其他平台上启用外接程序的功能。 进行此更改时，请务必在这些平台上测试功能。

> [!IMPORTANT]
> 清单不能指定`ExcelApiOnline 1.1`为激活要求。 不是在[Set 元素](../manifest/set.md)中使用的有效值。

## <a name="api-list"></a>API 列表

目前尚无仅联机的 Api。 重新检查，因为新功能将添加到 web 上的 Excel 中并受 Office JavaScript Api 支持。

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 参考文档](/javascript/api/excel?view=excel-js-online)
- [Excel JavaScript 预览 API](./excel-preview-apis.md)
- [Excel JavaScript API 要求集](./excel-api-requirement-sets.md)