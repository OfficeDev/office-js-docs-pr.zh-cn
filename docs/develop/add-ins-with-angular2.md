---
title: 使用 Angular 开发 Office 加载项
description: 使用 Angular 创建一Office外接程序作为单个页面应用程序。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: e0d30b7cb2f3d5489f5dae9e257c0cfc115a955e
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773501"
---
# <a name="develop-office-add-ins-with-angular"></a>使用 Angular 开发 Office 加载项

本文介绍了如何使用 Angular 2+ 将 Office 加载项创建为单页应用。

> [!NOTE]
> 根据自身体验，是否要参与有关使用 Angular 创建 Office 加载项的文章？ 你可以为本文[做贡献，GitHub](https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/develop/add-ins-with-angular2.md)提交问题来提供反馈。 [](https://github.com/OfficeDev/office-js-docs-pr/issues)

有关使用 Angular 框架生成的 Office 加载项示例，请参阅[使用 Angular 生成的 Word 样式检查加载项](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。

## <a name="install-the-typescript-type-definitions"></a>安装 TypeScript 类型定义

打开Node.js窗口，在命令行中输入以下内容。

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>启动代码必须位于 Office.initialize 内

在调用 JavaScript API Office、Word Excel的任何页面上，代码必须先为 属性分配 `Office.initialize` 方法。  (如果没有初始化代码，则方法正文可以只是空的" " 符号，但不得使属性 `{}` `Office.initialize` 保持未定义状态。 有关详细信息，请参阅初始化[你的](initialize-add-in.md)Office 外接程序 .) Office在初始化 JavaScript 库后立即Office此方法。

**Angular bootstrapping 代码必须在你分配到 `Office.initialize` 的方法内调用**，以确保 Office JavaScript 库首先进行了初始化。以下是演示如何执行该操作的简单示例。此代码应在项目的 main.ts 文件中。

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>使用 Angular 应用程序中的哈希位置策略

如果不指定哈希位置策略，则在应用程序中的路由间导航可能不起作用。可使用以下两种方法之一完成此操作：首先，可为应用模块中的位置策略指定提供程序，如以下示例中所示。它将进入 app.module.ts 文件。

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity

@NgModule({
  providers: [
    { provide: LocationStrategy, useClass: HashLocationStrategy },
    // Other providers suppressed
  ],
  // Other module properties suppressed
})
export class AppModule { }
```

如果在单独的路由模块中指定路由，则有指定哈希位置策略的替代方法。在路由模块的 .ts 文件中，向指定策略的 `forRoot` 函数传递配置对象。以下代码是一个示例。

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity

const routes: Routes = // route definitions go here

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
```

## <a name="use-the-office-dialog-api-with-angular"></a>将Office对话框 API 与 Angular

Office 加载项对话框 API 可使加载项打开非模态对话框中的页面，该页面可与主页面交换信息，这在任务窗格中是典型操作。

[displayDialogAsync](/javascript/api/office/office.ui) 方法采用指定应在对话框中打开的页面的 URL 的参数。外接程序可具有单独的 HTML 页面（与基本页不同）来传递此参数，或在 Angular 应用程序中传递路由的 URL。

要记住的重要一点是，如果传递路由，则该对话框将创建具有自身执行上下文的新窗口。 基本页及其所有初始化和引导代码将在新上下文中再次运行，且任何变量都将被设置为对话框中的初始值。 所以，此技术在对话框中启动了单页应用程序的第二个实例。 更改了对话框中的变量的代码不会更改同一变量的任务窗格版本。 同样，对话框具有其自己的会话存储 ([Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) ，任务窗格中的代码无法访问该存储。  

## <a name="trigger-the-ui-update"></a>触发 UI 更新

在 Angular 应用中，UI 有时不更新。这是因为部分代码在 Angular 区域外运行。解决方案将代码放在该区域中，如下面的示例所示。

```js
import { NgZone } from '@angular/core';

export class MyComponent {
  constructor(private zone: NgZone) { }

  myFunction() {
    this.zone.run(() => {
      // the codes that need update the UI
    });
  }
}
```

## <a name="use-observable"></a>使用可观察

Angular 使用 RxJS (Reactive Extensions for JavaScript)，而 RxJS 引入了 `Observable` 和 `Observer` 对象来实现异步处理。本部分简要介绍了如何使用 `Observables`；有关详细信息，请参阅官方 [RxJS](https://rxjs-dev.firebaseapp.com/) 文档。

`Observable` 在某种程度上类似一个 `Promise` 对象 - 它立即从异步调用返回，但它可能在以后才能进行解析。不过，`Promise` 是一个值（可以是一个数组对象），而 `Observable` 是对象数组（可能仅有一个成员）。这可使代码调用 `Observable` 对象上的 [数组方法](https://www.w3schools.com/jsref/jsref_obj_array.asp)，如 `concat`、`map` 和 `filter`。

### <a name="push-instead-of-pull"></a>推送而不是下拉

代码“拉取” `Promise` 对象（通过将其分配到变量），而 `Observable` 对象将其值“推送”到 *订阅* `Observable` 的对象。订阅服务器是 `Observer` 对象。推送体系结构的优势是，随着时间的推移，新成员可以添加到 `Observable` 数组。添加了新成员后，订阅 `Observable` 的所有 `Observer` 对象都将收到一条通知。

`Observer` 被配置为使用一个函数处理每个新对象（称为“next”对象）。（它还被配置为响应一个错误和一个完成通知。参阅下一部分的一个示例。）为此，与 `Promise` 对象相比，`Observable` 对象的使用范围更广。例如，除了从 AJAX 调用返回 `Observable`（即返回 `Promise` 的方式）以外，还可从事件处理程序返回 `Observable`，例如文本框的“已更改”事件处理程序。用户每次在框中输入文本时，所有已订阅的 `Observer` 对象将使用最新文本和/或输入的应用程序的当前状态立即做出反应。

### <a name="wait-until-all-asynchronous-calls-have-completed"></a>等待所有异步调用完成

如果想要确保仅当一组 `Promise` 对象的每个成员解析后运行撤消，请使用 `Promise.all()` 方法。

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
```

若要使用 `Observable` 对象执行同一操作，请使用 [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) 方法。  

```js
const source = Observable.forkJoin([x, y, z]);

const subscription = source.subscribe(
  x => {
    // TODO: Callback logic goes here
  },
  err => console.log('Error: ' + err),
  () => console.log('Completed')
);
```

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a>使用 Ahead-of-Time (AOT) 编译器编译 Angula r应用程序

应用程序性能是影响用户体验的最重要方面之一。 可以在构建时使用 Angular Ahead-of-Time (AOT) 编译器编译应用程序来优化 Angular 应用程序。 它可将所有源代码（HTML 模板和 TypeScript）转换为高效的 JavaScript 代码。 如果使用 AOT 编译器编译应用程序，则运行时不会发生其他编译，从而加快 HTML 模板的呈现和异步请求。 此外，应用程序的总体大小将减小，因为 Angular 编译器无需包含在可分发应用程序中。

若要使用 AOT 编译器，请将 `--aot` 添加到 `ng build` 或 `ng serve` 命令：

```command&nbsp;line
ng build --aot
ng serve --aot
```

> [!NOTE]
> 若要了解有关 Angular Ahead-of-Time (AOT) 编译器的详细信息，请参阅[官方指南](https://angular.io/guide/aot-compiler)。

## <a name="support-internet-explorer-if-youre-dynamically-loading-officejs"></a>如果你Internet Explorer加载内容，支持Office.js

根据 Windows 版本Office运行外接程序的桌面客户端，外接程序可能正在使用 Internet Explorer 11。  (有关详细信息，请参阅[Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Angular 使用的浏览器取决于一些 API，但这些 API 在 Windows 桌面客户端中嵌入的 IE 运行时中不起作用。 `window.history` 当这些 API 不起作用时，您的外接程序可能无法正常运行，例如，它可能会加载一个空白的任务窗格。 若要缓解此Office.js，请空值这些 API。 但是，如果你要动态加载 Office.js，AngularJS 可能会先加载Office.js。 在这种情况下，应该通过将以下代码添加到加载项的"index.htm`window.history` **l"页面来禁用这些** API。

```js
<script type="text/javascript">window.history.replaceState=null;window.history.pushState=null;</script>
```
