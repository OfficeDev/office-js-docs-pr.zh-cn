---
layout: LandingPage
ms.topic: landing-page
title: Office JavaScript API 参考文档
description: 了解 Office JavaScript API。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 5622146d9663881eea0a97cafa5e793aa0381932
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720741"
---
# <a name="api-reference-documentation"></a>API 参考文档

加载项可使用 Office JavaScript API 与 Office 主机应用程序中的对象进行交互。 

<ul>
    <li><b>主机特定的</b> API 提供了强类型对象，这种对象可用于与特定 Office 应用程序的本机对象进行交互。</li>
    <li><b>通用</b> API 可用于访问在多种类型的 Office 应用程序中都很常见的 UI、对话框和客户端设置等功能。</li>
</ul>

应尽可能使用主机特定的 API，并仅在主机特定的 API 不支持的情况中使用通用 API。 有关这两种 API 模型的更多详细信息，请参阅<a href="../overview/office-add-ins-fundamentals.md#api-models">构建 Office 加载项</a>。

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

<b>注意</b>：对于 Project，目前没有主机特定的 JavaScript API，因此需要使用通用 API 创建 Project 加载项。此外，对于 PowerPoint，主机特定的 API 的范围非常有限，因此主要使用通用 API 创建 PowerPoint 加载项。

<h2>开放 API 规范</h2>

在我们设计和开发新的 API 以用于 Office 外接程序时，我们将使它们适用于[开放 API 规范](openspec/openspec.md)页的反馈。了解管道中的新增功能，并提供您对我们的设计规范的宝贵意见。