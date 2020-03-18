---
ms.date: 07/10/2019
description: 使用 `OfficeRuntime.storage` 保存自定义函数中的状态。
title: 保存并共享自定义函数中的状态
localization_priority: Normal
ms.openlocfilehash: 8b55bfe61595b91f01c587282dc3f34887ce50fb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717199"
---
# <a name="save-and-share-state-in-custom-functions"></a>保存并共享自定义函数中的状态

使用 `OfficeRuntime.storage` 对象保存与加载项中的自定义函数或任务窗格相关的状态。 存储限制为每个域 10 MB（可以在多个加载项中共享）。 在 Windows 版 Excel 中，`storage` 对象是自定义函数运行时内的单独位置；但对于 Excel 网页版和 Mac 版 Excel，`storage` 对象与浏览器的 `localStorage` 相同。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

可以通过多种方式使用 `storage` 进行状态管理：

- 可以存储自定义函数的默认值，以便在你离线和无法触及网页资源时使用。
- 可以存储自定义函数值，以免额外调用网页资源。
- 可以保存自定义函数中的值。
- 可以存储任务窗格中的值。

以下代码示例演示了如何将项存储于 `storage` 中并检索它。

```js
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}
```

[GitHub 上的更详细代码示例](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)提供了将此信息传递到任务窗格的示例。

>[!NOTE]
> `storage` 对象取代了先前名为 `AsyncStorage` 的存储对象（现已启用）。 如果在当前的自定义函数代码中使用 `AsyncStorage` 对象，请将其更新为使用 `storage` 对象。

## <a name="next-steps"></a>后续步骤
了解如何[为自定义函数自动生成 JSON 元数据](custom-functions-json-autogeneration.md)。 

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
* [自定义函数调试](custom-functions-debugging.md)
