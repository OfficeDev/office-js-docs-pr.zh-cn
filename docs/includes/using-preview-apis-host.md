> [!NOTE]
> 预览 API 可能会发生变更，不适合在生产环境中使用。 我们建议你仅在测试和开发环境中试用它们。 不要在生产环境或业务关键型文档中使用预览 API。
>
> 若要使用预览 API：
>
> - 你必须从内容交付网络 Office 使用Office.js JavaScript API 库[的预览 (CDN) 。 ](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 用于 TypeScript 编译和 IntelliSense 的[类型定义文件](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)位于 CDN 和 [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) 中。 可以使用 `npm install --save-dev @types/office-js-preview` 来安装这些类型。
> - 可能需要加入 [Office 预览体验计划](https://insider.office.com)才能访问更新的 Office 版本。