<div align="center">

# officejj

<p>
  <strong>让 Java 里的 Word、Excel、PDF 处理，回到“能快点做完”的状态。</strong>
</p>

<p>
  <a href="https://search.maven.org/search?q=g:cn.javaex%20AND%20a:officejj">
    <img src="https://img.shields.io/badge/Maven-officejj-blue" alt="Maven">
  </a>
  <a href="https://doc.javaex.cn/officejj">
    <img src="https://img.shields.io/badge/Docs-Read%20the%20Guide-brightgreen" alt="Docs">
  </a>
  <a href="LICENSE">
    <img src="https://img.shields.io/badge/License-Apache%202.0-orange" alt="License">
  </a>
  <img src="https://img.shields.io/badge/Java-8%2B-success" alt="Java 8+">
</p>

<p>
  面向业务场景的 Office 工具包，帮助你更高效地处理 <strong>Word / Excel / PDF</strong>。
</p>

</div>

---

## 为什么你会想试试它？

很多 Java 项目里，文档处理最后都会慢慢变成这样：

- Excel 导出能跑，但代码越来越长
- Word 模板一复杂，维护起来就开始痛苦
- PDF 输出、字体、排版、下载响应，全是零碎问题
- 每个项目都在重复写一套自己的导出工具类

`officejj` 想做的，不是再给你一堆底层 API，
而是让这些常见需求更快落地、更像业务代码、更适合长期维护。

一句话概括：

> **少写样板代码，把精力放回业务本身。**

---

## 它适合哪些场景？

### Excel 导出
适合后台列表导出、报表导出、批量数据导出、模板化数据输出。

### Word 模板生成
适合通知书、证明、合同、申请单、报告等固定格式文档生成。

### PDF 输出
适合打印件、归档件、回执单、下载件、对外分发文档等场景。

---

## 你可能会喜欢的地方

### 1. 更贴近业务表达
很多时候我们真正关心的不是：
- 某个底层对象怎么创建
- 某个单元格样式怎么一层层设置

而是：
- 这一列怎么展示
- 这份模板里哪些字段要替换
- 这批数据怎么输出成用户想要的格式
- 这个文件最终怎么交付给前端下载

`officejj` 更偏向解决后面这类问题。

### 2. 更适合项目长期维护
文档处理最难的，从来不是第一次写出来，
而是几个月后还敢不敢改、同事接不接得住。

如果工具能让代码更集中、更规整、更容易复用，它就不仅仅是“能用”，而是“值得留在项目里”。

### 3. 一套思路覆盖 Word / Excel / PDF
很多业务系统不是只需要 Excel，
而是 Word、Excel、PDF 都会碰到。

统一思路，通常比项目里混用多套不同方案更舒服。

---

## 和原生 POI 的区别，更像什么？

`officejj` 不是要替代底层能力，
而是在很多常见业务场景里，帮你把底层复杂度“包起来”。

| 对比项 | 原生方式 | officejj |
|---|---|---|
| 上手体验 | 需要了解较多底层对象 | 更偏业务化使用 |
| 样板代码 | 通常较多 | 更倾向减少重复代码 |
| 常见导出场景 | 自己封装较多 | 更适合直接落地 |
| 团队维护 | 依赖个人经验 | 更容易形成统一写法 |
| Word / Excel / PDF 统一体验 | 往往分散处理 | 更适合统一思路 |

如果你的需求只是极简单的单表导出，原生方式当然也能做。  
但只要业务里文档处理越来越多，`officejj` 的优势通常会开始显现。

---

## 3 分钟快速开始

### 1）引入依赖

```xml
<dependency>
    <groupId>cn.javaex</groupId>
    <artifactId>officejj</artifactId>
    <version>6.1.1</version>
</dependency>
```

### 2）打开文档

- 文档地址：`https://doc.javaex.cn/officejj`
- 官网：`https://www.javaex.cn`
- Excel 异步导入：见 [EXCEL_ASYNC_IMPORT.md](EXCEL_ASYNC_IMPORT.md)

### 3）从一个最小需求开始试

最推荐你这样开始：

- 先做一个 Excel 列表导出
- 或先跑一个 Word 模板填充
- 或先做一个 PDF 下载输出

不要一开始就挑战最复杂的模板。  
先跑通一个最小案例，你会更快感受到它是否适合你的项目。

---

## 推荐的尝试顺序

如果你第一次接触 `officejj`，建议按这个顺序体验：

### 方案 A：先从 Excel 开始
最容易看见成果，也最容易快速接进已有后台项目。

### 方案 B：再试 Word 模板
如果你项目里有通知书、合同、证明、申请表这类固定格式文档，这一步通常会很有感觉。

### 方案 C：最后补 PDF 输出
当你已经有文档生成需求时，PDF 往往会自然出现。

这样的体验路径，通常比一上来研究全部功能更轻松。

---

## 一个真实的开发者心态

你未必需要“最强”的文档处理框架。  
大多数时候，你真正需要的是一个：

- 能快速接进项目
- 常见需求不用自己反复造轮子
- 代码风格尽量统一
- 团队成员更容易接手
- 做出来的功能能稳定交付

的工具。

如果这就是你现在想找的东西，`officejj` 很值得你打开文档跑一个例子。

---

## 什么时候你会明显感受到它的价值？

当你遇到下面这些需求时，感受会尤其明显：

- 同一个系统里同时存在 Word、Excel、PDF 处理
- 导出逻辑越来越多，工具类越堆越厚
- 文档样式、模板替换、动态内容开始变复杂
- 项目要交给别人维护，不想只靠“写的人自己懂”

这时候，一个更贴近业务场景的工具包，往往会比单纯堆底层代码更省心。


---

## 最后

如果你正在做一个有导入导出、模板文档、报表、打印件、归档件需求的 Java 项目，
与其继续在底层 API 上一层层堆工具类，
不如直接给 `officejj` 一个机会。

**先跑一个最小案例。**  
很多时候，它是不是适合你的项目，十分钟内就会有答案。

#### 示例
word
![输入图片说明](https://webimg.javaex.cn/FtNOVk2ubmhackvFPVEzhSuySO68)

excel
内置多线程导出方法，导出100W条数据，仅需21秒
![输入图片说明](https://webimg.javaex.cn/Fs-ROZubteH2iBp2pmWt2P2B0OJL)

pdf

![输入图片说明](https://webimg.javaex.cn/Fiu1c4wKumBJ9Jfy3Mxym05FCXb6)
