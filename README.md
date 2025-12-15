# html-to-pptx-autofixer          
HTML转换为PPTX的同时，提供自动修复HTML文件的功能

## 功能特点

### 1. **智能HTML修复**
- 根据HTML文件中存在的错误信息，自动修复常见问题。
- 修复完成后，会保存修复后的HTML文件（支持备份原始文件）。

### 2. **支持的修复类型**
- **文本元素边框修复器 (TextElementBorderFixer)**:
  - 修复错误：文本元素（如`<h1>`、`<p>`等）包含不支持的样式（边框、背景、阴影）。
  - 修复方法：将这些样式移至外层的`<div>`容器中。
  
- **未包裹文本修复器 (UnwrappedTextFixer)**:
  - 修复错误：`<div>`元素中包含未被标签包裹的文本。
  - 修复方法：将未包裹的文本用适当的HTML标签（如`<p>`、`<h1>`-`<h6>`、`<ul>`、`<ol>`等）包裹。

- **CSS渐变修复器 (CssGradientFixer)**:
  - 修复错误：CSS中的线性渐变、径向渐变和背景图片不被支持。
  - 修复方法：将CSS渐变替换为单色背景，将背景图片替换为默认背景色。

### 3. **修复过程**
- 根据错误信息，自动选择合适的修复器。
- 修复后将HTML文件保存到原路径。
- 支持备份原始HTML文件，避免文件丢失。

---

## 示例

### 示例1：修复带有边框的文本元素
**原始HTML代码**：
```html
<h1 class="title" style="border: 1px solid red;">标题</h1>
```

**修复后HTML代码**：
```html
<div class="title">
  <h1>标题</h1>
</div>
<style>
  .title {
    border: 1px solid red;
  }
  .title h1 {
    /* 原有文字样式保留 */
  }
</style>
```

### 示例2：修复未包裹的文本
**原始HTML代码**：
```html
<div class="content">
  这是未包裹的文本
</div>
```

**修复后HTML代码**：
```html
<div class="content">
  <p>这是未包裹的文本</p>
</div>
```

### 示例3：修复CSS渐变
**原始CSS代码**：
```css
.background {
  background: linear-gradient(to right, red, blue);
}
```

**修复后CSS代码**：
```css
.background {
  background: red;
}
```

---
---

## 安装方法

1. 克隆项目到本地：
   ```bash
   git clone <仓库地址>
   cd <项目目录>
   ```

2. 安装依赖：
   ```bash
   npm install
   ```

---

## 项目限制

1. **不支持的HTML特性**：
   - 渐变背景将被替换为单色背景。
   - 背景图片将被替换为默认的浅灰色背景。
   - 某些复杂的嵌套样式可能需要手动修复。

2. **依赖项**：
   - 本工具依赖`jsdom`库进行HTML的DOM解析，请确保已安装依赖。

---
