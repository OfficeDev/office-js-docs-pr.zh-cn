---
title: 使用 Angular 开发 Office 加载项
description: 使用 Angular 创建一Office外接程序作为单个页面应用程序。
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: e12f3e2d4733613fb542cf2be4e0ff6648ab8475
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350083"
---
# <a name="develop-office-add-ins-with-angular"></a><span data-ttu-id="6d9b2-103">使用 Angular 开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6d9b2-103">Develop Office Add-ins with Angular</span></span>

<span data-ttu-id="6d9b2-104">本文介绍了如何使用 Angular 2+ 将 Office 加载项创建为单页应用。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-104">This article provides guidance for using Angular 2+ to create an Office Add-in as a single page application.</span></span>

> [!NOTE]
> <span data-ttu-id="6d9b2-105">根据自身体验，是否要参与有关使用 Angular 创建 Office 加载项的文章？</span><span class="sxs-lookup"><span data-stu-id="6d9b2-105">Do you have something to contribute based on your experience using Angular to create Office Add-ins?</span></span> <span data-ttu-id="6d9b2-106">你可以为本文[做贡献，GitHub](https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/develop/add-ins-with-angular2.md)提交问题来提供反馈。 [](https://github.com/OfficeDev/office-js-docs-pr/issues)</span><span class="sxs-lookup"><span data-stu-id="6d9b2-106">You can contribute to [this article in GitHub](https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/develop/add-ins-with-angular2.md) or provide your feedback by submitting an [issue](https://github.com/OfficeDev/office-js-docs-pr/issues) in the repo.</span></span>

<span data-ttu-id="6d9b2-107">有关使用 Angular 框架生成的 Office 加载项示例，请参阅[使用 Angular 生成的 Word 样式检查加载项](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-107">For an Office Add-ins sample that's built using the Angular framework, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span></span>

## <a name="install-the-typescript-type-definitions"></a><span data-ttu-id="6d9b2-108">安装 TypeScript 类型定义</span><span class="sxs-lookup"><span data-stu-id="6d9b2-108">Install the TypeScript type definitions</span></span>

<span data-ttu-id="6d9b2-109">打开Node.js窗口，在命令行中输入以下内容。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-109">Open a Node.js window and enter the following at the command line.</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a><span data-ttu-id="6d9b2-110">启动代码必须位于 Office.initialize 内</span><span class="sxs-lookup"><span data-stu-id="6d9b2-110">Bootstrapping must be inside Office.initialize</span></span>

<span data-ttu-id="6d9b2-111">在调用 JavaScript API Office、Word Excel的任何页面上，代码必须先为 属性分配 `Office.initialize` 方法。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-111">On any page that calls the Office, Word, or Excel JavaScript APIs, your code must first assign a method to the `Office.initialize` property.</span></span> <span data-ttu-id="6d9b2-112"> (如果没有初始化代码，则方法正文可以只是空的" " 符号，但不得使属性 `{}` `Office.initialize` 保持未定义状态。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-112">(If you have no initialization code, the method body can be just empty "`{}`" symbols, but you must not leave the `Office.initialize` property undefined.</span></span> <span data-ttu-id="6d9b2-113">有关详细信息，请参阅初始化[你的](initialize-add-in.md)Office 外接程序 .) Office在初始化 JavaScript 库后立即Office此方法。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-113">For details, see [Initialize your Office Add-in](initialize-add-in.md).) Office calls this method immediately after it has initialized the Office JavaScript libraries.</span></span>

<span data-ttu-id="6d9b2-p103">**Angular bootstrapping 代码必须在你分配到 `Office.initialize` 的方法内调用**，以确保 Office JavaScript 库首先进行了初始化。以下是演示如何执行该操作的简单示例。此代码应在项目的 main.ts 文件中。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p103">**Your Angular bootstrapping code must be called inside the method that you assign to `Office.initialize`** to ensure that the Office JavaScript libraries have initialized first. The following is a simple example that shows how to do this. This code should be in the main.ts file of the project.</span></span>

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a><span data-ttu-id="6d9b2-117">使用 Angular 应用程序中的哈希位置策略</span><span class="sxs-lookup"><span data-stu-id="6d9b2-117">Use the hash location strategy in the Angular application</span></span>

<span data-ttu-id="6d9b2-p104">如果不指定哈希位置策略，则在应用程序中的路由间导航可能不起作用。可使用以下两种方法之一完成此操作：首先，可为应用模块中的位置策略指定提供程序，如以下示例中所示。它将进入 app.module.ts 文件。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p104">Navigating between routes in the application might not work if you don't specify the hash location strategy. You can do this in one of two ways. First, you can specify a provider for the location strategy in your app module, as shown in the following example. It goes into the app.module.ts file.</span></span>

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

<span data-ttu-id="6d9b2-p105">如果在单独的路由模块中指定路由，则有指定哈希位置策略的替代方法。在路由模块的 .ts 文件中，向指定策略的 `forRoot` 函数传递配置对象。以下代码是一个示例。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p105">If you define your routes in a separate routing module, there is an alternative way to specify the hash location strategy. In your routing module's .ts file, pass a configuration object to the `forRoot` function that specifies the strategy. The following code is an example.</span></span>

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

## <a name="using-the-office-dialog-api-with-angular"></a><span data-ttu-id="6d9b2-125">将 Office 对话框 API 与 Angular 结合使用</span><span class="sxs-lookup"><span data-stu-id="6d9b2-125">Using the Office dialog API with Angular</span></span>

<span data-ttu-id="6d9b2-126">Office 加载项对话框 API 可使加载项打开非模态对话框中的页面，该页面可与主页面交换信息，这在任务窗格中是典型操作。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-126">The Office Add-in dialog API enables your add-in to open a page in a nonmodal dialog box that can exchange information with the main page, which is typically in a task pane.</span></span>

<span data-ttu-id="6d9b2-p106">[displayDialogAsync](/javascript/api/office/office.ui) 方法采用指定应在对话框中打开的页面的 URL 的参数。外接程序可具有单独的 HTML 页面（与基本页不同）来传递此参数，或在 Angular 应用程序中传递路由的 URL。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p106">The [displayDialogAsync](/javascript/api/office/office.ui) method takes a parameter that specifies the URL of the page that should open in the dialog box. Your add-in can have a separate HTML page (different from the base page) to pass to this parameter, or you can pass the URL of a route in your Angular application.</span></span>

<span data-ttu-id="6d9b2-129">要记住的重要一点是，如果传递路由，则该对话框将创建具有自身执行上下文的新窗口。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-129">It is important to remember, if you pass a route, that the dialog box creates a new window with its own execution context.</span></span> <span data-ttu-id="6d9b2-130">基本页及其所有初始化和引导代码将在新上下文中再次运行，且任何变量都将被设置为对话框中的初始值。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-130">Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box.</span></span> <span data-ttu-id="6d9b2-131">所以，此技术在对话框中启动了单页应用程序的第二个实例。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-131">So this technique launches a second instance of your single page application in the dialog box.</span></span> <span data-ttu-id="6d9b2-132">更改了对话框中的变量的代码不会更改同一变量的任务窗格版本。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-132">Code that changes variables in the dialog box does not change the task pane version of the same variables.</span></span> <span data-ttu-id="6d9b2-133">同样，对话框具有其自己的会话存储 ([Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) ，任务窗格中的代码无法访问该存储。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-133">Similarly, the dialog box has its own session storage (the [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) property), which is not accessible from code in the task pane.</span></span>  

## <a name="trigger-the-ui-update"></a><span data-ttu-id="6d9b2-134">触发 UI 更新</span><span class="sxs-lookup"><span data-stu-id="6d9b2-134">Trigger the UI update</span></span>

<span data-ttu-id="6d9b2-p108">在 Angular 应用中，UI 有时不更新。这是因为部分代码在 Angular 区域外运行。解决方案将代码放在该区域中，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p108">In an Angular app, the UI sometimes does not update. This is because that part of the code runs out of the Angular zone. The solution is to put the code in the zone, as shown in the following example.</span></span>

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

## <a name="using-observable"></a><span data-ttu-id="6d9b2-138">使用 Observable 对象</span><span class="sxs-lookup"><span data-stu-id="6d9b2-138">Using Observable</span></span>

<span data-ttu-id="6d9b2-p109">Angular 使用 RxJS (Reactive Extensions for JavaScript)，而 RxJS 引入了 `Observable` 和 `Observer` 对象来实现异步处理。本部分简要介绍了如何使用 `Observables`；有关详细信息，请参阅官方 [RxJS](https://rxjs-dev.firebaseapp.com/) 文档。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p109">Angular uses RxJS (Reactive Extensions for JavaScript), and RxJS introduces `Observable` and `Observer` objects to implement asynchronous processing. This section provides a brief introduction to using `Observables`; for more detailed information, see the official [RxJS](https://rxjs-dev.firebaseapp.com/) documentation.</span></span>

<span data-ttu-id="6d9b2-p110">`Observable` 在某种程度上类似一个 `Promise` 对象 - 它立即从异步调用返回，但它可能在以后才能进行解析。不过，`Promise` 是一个值（可以是一个数组对象），而 `Observable` 是对象数组（可能仅有一个成员）。这可使代码调用 `Observable` 对象上的 [数组方法](https://www.w3schools.com/jsref/jsref_obj_array.asp)，如 `concat`、`map` 和 `filter`。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p110">An `Observable` is like a `Promise` object in some ways - it is returned immediately from an asynchronous call, but it might not resolve until some time later. However, while a `Promise` is a single value (which can be an array object), an `Observable` is an array of objects (possibly with only a single member). This enables code to call [array methods](https://www.w3schools.com/jsref/jsref_obj_array.asp), such as `concat`, `map`, and `filter`, on `Observable` objects.</span></span>

### <a name="pushing-instead-of-pulling"></a><span data-ttu-id="6d9b2-144">使用“推送”代替“拉取”</span><span class="sxs-lookup"><span data-stu-id="6d9b2-144">Pushing instead of pulling</span></span>

<span data-ttu-id="6d9b2-p111">代码“拉取” `Promise` 对象（通过将其分配到变量），而 `Observable` 对象将其值“推送”到 *订阅* `Observable` 的对象。订阅服务器是 `Observer` 对象。推送体系结构的优势是，随着时间的推移，新成员可以添加到 `Observable` 数组。添加了新成员后，订阅 `Observable` 的所有 `Observer` 对象都将收到一条通知。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p111">Your code "pulls" `Promise` objects by assigning them to variables, but `Observable` objects "push" their values to objects that *subscribe* to the `Observable`. The subscribers are `Observer` objects. The benefit of the push architecture is that new members can be added to the `Observable` array over time. When a new member is added, all the `Observer` objects that subscribe to the `Observable` receive a notification.</span></span>

<span data-ttu-id="6d9b2-p112">`Observer` 被配置为使用一个函数处理每个新对象（称为“next”对象）。（它还被配置为响应一个错误和一个完成通知。参阅下一部分的一个示例。）为此，与 `Promise` 对象相比，`Observable` 对象的使用范围更广。例如，除了从 AJAX 调用返回 `Observable`（即返回 `Promise` 的方式）以外，还可从事件处理程序返回 `Observable`，例如文本框的“已更改”事件处理程序。用户每次在框中输入文本时，所有已订阅的 `Observer` 对象将使用最新文本和/或输入的应用程序的当前状态立即做出反应。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-p112">The `Observer` is configured to process each new object (called the "next" object) with a function. (It is also configured to respond to an error and a completion notification. See the next section for an example.) For this reason, `Observable` objects can be used in a wider range of scenarios than `Promise` objects. For example, in addition to returning an `Observable` from an AJAX call, the way you can return a `Promise`, an `Observable` can be returned from an event handler, such as the "changed" event handler for a text box. Each time a user enters text in the box, all the subscribed `Observer` objects react immediately using the latest text and/or the current state of the application as input.</span></span>

### <a name="waiting-until-all-asynchronous-calls-have-completed"></a><span data-ttu-id="6d9b2-154">等待所有异步调用完成</span><span class="sxs-lookup"><span data-stu-id="6d9b2-154">Waiting until all asynchronous calls have completed</span></span>

<span data-ttu-id="6d9b2-155">如果想要确保仅当一组 `Promise` 对象的每个成员解析后运行撤消，请使用 `Promise.all()` 方法。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-155">When you want to ensure that a callback only runs when every member of a set of `Promise` objects has resolved, use the `Promise.all()` method.</span></span>

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
```

<span data-ttu-id="6d9b2-156">若要使用 `Observable` 对象执行同一操作，请使用 [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) 方法。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-156">To do the same thing with an `Observable` object, you use the [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) method.</span></span>  

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

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a><span data-ttu-id="6d9b2-157">使用 Ahead-of-Time (AOT) 编译器编译 Angula r应用程序</span><span class="sxs-lookup"><span data-stu-id="6d9b2-157">Compile the Angular application using the Ahead-of-Time (AOT) compiler</span></span>

<span data-ttu-id="6d9b2-158">应用程序性能是影响用户体验的最重要方面之一。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-158">Application performance is one of the most important aspects of user experience.</span></span> <span data-ttu-id="6d9b2-159">可以在构建时使用 Angular Ahead-of-Time (AOT) 编译器编译应用程序来优化 Angular 应用程序。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-159">An Angular application can be optimized by using the Angular Ahead-of-Time (AOT) compiler to compile the app at build time.</span></span> <span data-ttu-id="6d9b2-160">它可将所有源代码（HTML 模板和 TypeScript）转换为高效的 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-160">It converts all source code (HTML templates and TypeScript) into efficient JavaScript code.</span></span> <span data-ttu-id="6d9b2-161">如果使用 AOT 编译器编译应用程序，则运行时不会发生其他编译，从而加快 HTML 模板的呈现和异步请求。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-161">If you compile your app with the AOT compiler, no additional compilation will occur at runtime, which results in faster rendering and faster asynchronous requests for HTML templates.</span></span> <span data-ttu-id="6d9b2-162">此外，应用程序的总体大小将减小，因为 Angular 编译器无需包含在可分发应用程序中。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-162">Additionally, the overall application size will be reduced, because the Angular compiler won't need to be included in the application distributable.</span></span>

<span data-ttu-id="6d9b2-163">若要使用 AOT 编译器，请将 `--aot` 添加到 `ng build` 或 `ng serve` 命令：</span><span class="sxs-lookup"><span data-stu-id="6d9b2-163">To use the AOT compiler, add `--aot` to the `ng build` or `ng serve` command:</span></span>

```command&nbsp;line
ng build --aot
ng serve --aot
```

> [!NOTE]
> <span data-ttu-id="6d9b2-164">若要了解有关 Angular Ahead-of-Time (AOT) 编译器的详细信息，请参阅[官方指南](https://angular.io/guide/aot-compiler)。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-164">To learn more about the Angular Ahead-of-Time (AOT) compiler, see the [official guide](https://angular.io/guide/aot-compiler).</span></span>

## <a name="support-internet-explorer-if-youre-dynamically-loading-officejs"></a><span data-ttu-id="6d9b2-165">如果你Internet Explorer加载内容，支持Office.js</span><span class="sxs-lookup"><span data-stu-id="6d9b2-165">Support Internet Explorer if you're dynamically loading Office.js</span></span>

<span data-ttu-id="6d9b2-166">根据 Windows 版本Office运行外接程序的桌面客户端，外接程序可能正在使用 Internet Explorer 11。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-166">Based on the Windows version and the Office desktop client where your add-in is running, your add-in may be using Internet Explorer 11.</span></span> <span data-ttu-id="6d9b2-167"> (有关详细信息，请参阅[Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Angular 使用的浏览器取决于一些 API，但这些 API 在 Windows 桌面客户端中嵌入的 IE 运行时中不起作用。 `window.history`</span><span class="sxs-lookup"><span data-stu-id="6d9b2-167">(For more details, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).) Angular depends on a few `window.history` APIs but these APIs don't work in the IE runtime embedded in Windows desktop clients.</span></span> <span data-ttu-id="6d9b2-168">当这些 API 不起作用时，您的外接程序可能无法正常运行，例如，它可能会加载一个空白的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-168">When these APIs don't work, your add-in may not work properly, for example, it may load a blank task pane.</span></span> <span data-ttu-id="6d9b2-169">若要缓解此Office.js，请空值这些 API。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-169">To mitigate this, Office.js nullifies those APIs.</span></span> <span data-ttu-id="6d9b2-170">但是，如果你要动态加载 Office.js，AngularJS 可能会先加载Office.js。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-170">However, if you're dynamically loading Office.js, AngularJS may load before Office.js.</span></span> <span data-ttu-id="6d9b2-171">在这种情况下，应该通过将以下代码添加到加载项的"index.htm`window.history` **l"页面来禁用这些** API。</span><span class="sxs-lookup"><span data-stu-id="6d9b2-171">In that case, you should disable the `window.history` APIs by adding the following code to your add-in's **index.html** page.</span></span>

```js
<script type="text/javascript">window.history.replaceState=null;window.history.pushState=null;</script>
```
