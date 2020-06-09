> [!NOTE]
> 预览 API 可能会发生变更，不适合在生产环境中使用。 我们建议你仅在测试和开发环境中试用它们。 不要在生产环境或业务关键型文档中使用预览 API。
>
> 使用预览 Api 的步骤：
>
> - 您必须在 CDN 上引用**beta**库（ https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 。 在 CDN 和[jquery.typescript.definitelytyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)中找到 TypeScript 编译和智能感知的[类型定义文件](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。 您可以使用安装这些类型 `npm install --save-dev @types/office-js-preview` 。
> - 你可能需要加入[Office 预览体验成员计划](https://insider.office.com)，才能访问最新的 office 版本。