---
title: 清单文件中的 Hosts 元素
description: 指定将在其中激活 Office 外接程序的 Office 客户端应用程序。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 2684753fc32a295d7e177ef3bf668c194458128e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149510"
---
# <a name="hosts-element"></a>Hosts 元素

指定将在其中激活 Office 外接程序的 Office 客户端应用程序。包含 **Host** 元素及其设置的集合。 

当该元素被包括在 [VersionOverrides](versionoverrides.md)(#versionoverrides) 节点中时，它将替代清单中父级部分中的 **Hosts** 元素。 

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  是   |  说明主机及其设置。 |
