---
layout: LandingPage
ms.topic: landing-page
title: Office JavaScript API 参考文档
description: 了解 Office JavaScript API。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 46f5ae8a073fe45a6f6ac7f873f9adeb37fbab15
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294281"
---
# <a name="api-reference-documentation"></a>API 参考文档

加载项可使用 Office JavaScript API 与 Office 客户端应用程序中的对象进行交互。 

<ul>
    <li><b>应用程序特定的</b> API 提供了强类型对象，它可用于与特定 Office 应用程序的本机对象进行交互。</li>
    <li><b>通用</b> API 可用于访问在多种类型的 Office 应用程序中都很常见的 UI、对话框和客户端设置等功能。</li>
</ul>

应尽可能使用应用程序特定的 API，并仅在应用程序特定的 API 不支持的情况中使用通用 API。 有关这两种 API 模型的更多详细信息，请参阅<a href="../overview/office-add-ins-fundamentals.md#api-models">构建 Office 加载项</a>。

<h2>API 参考</h2>

<ul class="panelContent cardsF cols cols3">
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/excel"><img src="../images/index/logo-excel.svg" alt="Excel API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Excel API 参考</h3>
                        <p><a href="/javascript/api/excel">用于构建 Excel 加载项的 JavaScript API。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Outlook API 参考</h3>
                        <p><a href="/javascript/api/outlook">用于构建 Outlook 加载项的 JavaScript API。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/word"><img src="../images/index/logo-word.svg" alt="Word API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Word API 参考</h3>
                        <p><a href="/javascript/api/word">用于构建 Word 加载项的 JavaScript API。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>PowerPoint API 参考</h3>
                        <p><a href="/javascript/api/powerpoint">用于构建 PowerPoint 加载项的 JavaScript API。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>OneNote API 参考</h3>
                        <p><a href="/javascript/api/onenote">用于构建 OneNote 加载项的 JavaScript API。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/office"><img src="../images/index-landing-page/i_code-blocks.svg" alt="reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>通用 API 参考</h3>
                        <p><a href="/javascript/api/office">可由任意 Office 加载项使用的 JavaScript API。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
</ul>

<b>注意</b>：对于 Project，目前没有应用程序特定的 JavaScript API，因此需要使用通用 API 创建 Project 加载项。此外，对于 PowerPoint，应用程序特定的 API 的范围非常有限，因此主要使用通用 API 创建 PowerPoint 加载项。
