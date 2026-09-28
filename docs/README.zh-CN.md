# HTML 转可编辑 PPTX

这个转换器把浏览器渲染的 HTML 幻灯片导出为 PPTX。文字尽量写成原生文本框；复杂 CSS 装饰在必要时以局部图片保真。导出报告会标明降级情况。

## 先看效果

| Chrome 中的 HTML | LibreOffice 渲染的导出 PPTX |
| --- | --- |
| ![原始 HTML](media/gradient-slide-html.png) | ![导出 PPTX](media/gradient-slide-pptx.png) |

[下载可编辑 PPTX](media/gradient-slide.pptx) · [看 9 秒文字修改演示](media/editability-proof.mp4)

演示片通过修改 PPTX 文件里的原生文字再重新渲染来验证可编辑性，**并非 PowerPoint 操作界面录像**。示例页有 14 个可编辑文字框；两处复杂装饰成为局部图片。

## 对比其他工具

![两页对比：HTML、自研、dom-to-pptx](media/benchmark-comparison.png)

每行从左到右是原 HTML、自研导出、`dom-to-pptx 2.1.2`。深色渐变页自研更接近原图；浅色简洁页对照工具更接近原图。[看基准测试说明](../benchmarks/README.md)。

[下载修改文字后的 PPTX](media/gradient-slide-edited.pptx) · [查看验证方法](EDITABILITY.md)

## 立即试用

需要 Node.js 18+ 和 Chrome/Chromium。

从 GitHub 一条命令安装：

```bash
npm install github:H777000/html-to-editable-pptx
```

导出一个 HTML 文件：

```bash
npx html-to-editable-pptx slide.html --output slide.pptx --chrome /usr/bin/google-chrome
```

从源码运行浏览器演示：

```bash
git clone https://github.com/H777000/html-to-editable-pptx.git
cd html-to-editable-pptx
npm ci
npm run demo
```

打开 `http://127.0.0.1:4173`，点击 **Download PPTX**。命令行导出示例：

```bash
node bin/convert.cjs examples/gradient-slide.html --output examples/gradient-slide.pptx --chrome /usr/bin/google-chrome
```

如果 Chrome 路径不同，改用本机路径或设置 `CHROME_PATH`。

## 结果与限制

示例页导出 14 个可编辑文字框、5 个映射的视觉元素、2 个局部图片对象，没有触发整页图片兜底。复杂渐变、模糊、裁剪、伪元素等可能变成图片；文字提取失败时也可能整页退回图片。输出报告会标记这些情况。

可运行的示例和指标在 [benchmarks](../benchmarks/README.md)。示例 HTML 均为这个仓库单独编写。

项目采用 [MIT 许可证](../LICENSE)。`package.json` 中的 `private` 仅用于避免误发 npm，不影响 GitHub 源代码的使用。
