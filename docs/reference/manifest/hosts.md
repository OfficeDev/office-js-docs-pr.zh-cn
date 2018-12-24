---
title: 清单文件中的 Hosts 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 59010c0f6c0d14d8721856f81def11540db28704
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433409"
---
# <a name="hosts-element"></a>Hosts 元素

指定将在其中激活 Office 外接程序的 Office 客户端应用程序。包含 **Host** 元素及其设置的集合。 

当该元素被包括在 [VersionOverrides](versionoverrides.md)(#versionoverrides) 节点中时，它将替代清单中父级部分中的 **Hosts** 元素。 

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  是   |  说明主机及其设置。 |
