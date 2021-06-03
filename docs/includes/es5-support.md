对于 Office 和[Windows](../concepts/browsers-used-by-office-web-add-ins.md)的一些版本，运行外接程序的 JavaScript 引擎由 Internet Explorer。 Internet Explorer引擎不支持 ES5 之后版本的 JavaScript。 这意味着，如果不进行特殊处理，外接程序所处理的 JavaScript 文件将无法使用在 ES5 之后添加到语言的语法、类型或方法。 这并不意味着您必须使用 ES5语法编写。 还有其他两个选项：

- 在 [ECMAScript 2015](https://www.w3schools.com/Js/js_es6.asp) (（也称为 ES6) 或更高版本 JavaScript）中编写代码，或在 TypeScript 中编写代码，然后使用编译器（如 [#A0](https://babeljs.io/) 或 [tsc](https://www.typescriptlang.org/index.html)）将代码编译为 ES5 JavaScript。
- 在 ECMAScript 2015 或更高版本的 JavaScript[](https://en.wikipedia.org/wiki/Polyfill_(programming))中编写，但也加载填充库（如[core-js，](https://github.com/zloirock/core-js)它使 IE 能够运行代码）。

有关这些选项的详细信息，请参阅 Support [Internet Explorer 11](../develop/support-ie-11.md)。
