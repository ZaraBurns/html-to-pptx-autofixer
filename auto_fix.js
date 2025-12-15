/**
 * HTML自动修复脚本
 * 根据html2pptx转换错误信息，智能修复HTML文件
 */

const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");

/**
 * 错误修复器基类
 */
class ErrorFixer {
  constructor(errorMessage, htmlPath) {
    this.errorMessage = errorMessage;
    this.htmlPath = htmlPath;
    this.htmlContent = fs.readFileSync(htmlPath, "utf-8");
    this.dom = new JSDOM(this.htmlContent);
    this.document = this.dom.window.document;
    this.fixed = false;
    this.fixDescription = "";
  }

  /**
   * 判断是否能修复此错误
   */
  canFix() {
    return false;
  }

  /**
   * 执行修复
   */
  fix() {
    throw new Error("Subclass must implement fix() method");
  }

  /**
   * 保存修复后的HTML
   */
  save(backupOriginal = true) {
    if (!this.fixed) {
      return false;
    }

    // 备份原文件
    if (backupOriginal) {
      const backupPath = this.htmlPath + ".backup";
      if (!fs.existsSync(backupPath)) {
        fs.copyFileSync(this.htmlPath, backupPath);
        console.log(`  📋 已备份原文件: ${path.basename(backupPath)}`);
      }
    }

    // 保存修复后的内容
    const fixedHtml = this.dom.serialize();
    fs.writeFileSync(this.htmlPath, fixedHtml, "utf-8");
    console.log(`  ✓ 已保存修复: ${this.fixDescription}`);
    return true;
  }
}

/**
 * 文本元素边框修复器
 * 处理错误: 文本元素 <h1> 存在 边框。仅 <div> 元素支持背景、边框和阴影，文本元素不支持。
 */
class TextElementBorderFixer extends ErrorFixer {
  canFix() {
    // 匹配错误信息: 文本元素 <xxx> 存在 (边框|背景|阴影)
    // 支持多个不同标签的错误
    const matches = [
      ...this.errorMessage.matchAll(/文本元素 <(\w+)> 存在 (边框|背景|阴影)/g),
    ];
    if (matches.length > 0) {
      // 提取所有需要修复的标签名和样式类型
      this.tagNames = [...new Set(matches.map((match) => match[1]))]; // 去重
      this.styleType = matches[0][2]; // 使用第一个错误的样式类型
      return true;
    }
    return false;
  }

  fix() {
    let totalFixCount = 0;
    const fixedTags = [];

    // 处理所有需要修复的标签类型
    this.tagNames.forEach((tagName) => {
      // 查找所有该类型的文本元素
      const elements = this.document.querySelectorAll(tagName);
      let fixCount = 0;

      elements.forEach((element) => {
        const className = element.className;
        if (!className) return;

        // 检查是否有需要分离的样式
        const styleElement = this.findStyleForClass(className);
        if (!styleElement) return;

        const { textStyles, containerStyles } = this.separateStyles(
          styleElement.textContent,
          className
        );

        // 如果没有容器样式（边框、背景、阴影等），则不需要修复
        if (!containerStyles.trim()) return;

        // 1. 更新CSS样式
        this.updateStyles(
          styleElement,
          className,
          textStyles,
          containerStyles,
          tagName
        );

        // 2. 包装HTML元素
        this.wrapElement(element, className, tagName);

        fixCount++;
      });

      if (fixCount > 0) {
        fixedTags.push(`${fixCount}个<${tagName}>`);
        totalFixCount += fixCount;
      }
    });

    if (totalFixCount > 0) {
      this.fixed = true;
      this.fixDescription = `将${fixedTags.join("、")}元素的${
        this.styleType
      }样式移至外层<div>`;
    }

    return this.fixed;
  }

  /**
   * 查找样式表中的类定义
   */
  findStyleForClass(className) {
    const styleElements = this.document.querySelectorAll("style");
    for (const styleEl of styleElements) {
      const content = styleEl.textContent;
      if (content.includes(`.${className}`)) {
        return styleEl;
      }
    }
    return null;
  }

  /**
   * 分离文本样式和容器样式
   */
  separateStyles(cssContent, className) {
    // 容器样式属性（需要移到外层div）
    const containerStyleProps = [
      "background",
      "background-color",
      "background-image",
      "background-size",
      "background-position",
      "background-repeat",
      "background-attachment",
      "border",
      "border-top",
      "border-right",
      "border-bottom",
      "border-left",
      "border-width",
      "border-style",
      "border-color",
      "border-radius",
      "box-shadow",
    ];

    // 提取.className的样式规则
    const classRegex = new RegExp(`\\.${className}\\s*\\{([^}]+)\\}`, "s");
    const match = cssContent.match(classRegex);

    if (!match) {
      return { textStyles: "", containerStyles: "" };
    }

    const styleBlock = match[1];
    const styleLines = styleBlock
      .split(";")
      .map((s) => s.trim())
      .filter((s) => s);

    const textStyleLines = [];
    const containerStyleLines = [];

    styleLines.forEach((line) => {
      const prop = line.split(":")[0].trim();
      const isContainerStyle = containerStyleProps.some(
        (cp) => prop === cp || prop.startsWith(cp + "-")
      );

      if (isContainerStyle) {
        containerStyleLines.push(line);
      } else {
        textStyleLines.push(line);
      }
    });

    const textStyles = textStyleLines.join(";\n            ");
    const containerStyles = containerStyleLines.join(";\n            ");

    return { textStyles, containerStyles };
  }

  /**
   * 更新CSS样式
   */
  updateStyles(styleElement, className, textStyles, containerStyles, tagName) {
    const oldContent = styleElement.textContent;
    const classRegex = new RegExp(`(\\.${className})\\s*\\{[^}]+\\}`, "s");

    // 构建新的样式规则
    let newRules = "";
    if (containerStyles) {
      newRules += `        .${className} {\n            ${containerStyles};\n        }\n`;
    }
    if (textStyles) {
      newRules += `        .${className} ${tagName} {\n            ${textStyles};\n        }`;
    }

    const newContent = oldContent.replace(classRegex, newRules);
    styleElement.textContent = newContent;
  }

  /**
   * 用div包装元素
   */
  wrapElement(element, className, tagName) {
    // 创建新的div容器
    const wrapper = this.document.createElement("div");
    wrapper.className = className;

    // 复制元素内容到新元素
    const newElement = this.document.createElement(tagName);
    newElement.innerHTML = element.innerHTML;

    // 复制其他属性（除了class）
    Array.from(element.attributes).forEach((attr) => {
      if (attr.name !== "class") {
        newElement.setAttribute(attr.name, attr.value);
      }
    });

    // 组装结构
    wrapper.appendChild(newElement);

    // 替换原元素
    element.parentNode.replaceChild(wrapper, element);
  }
}

/**
 * DIV 未包裹文本修复器
 * 处理错误: DIV 元素包含未包裹文本"xxx"。所有文本必须用 <p>、<h1>-<h6>、<ul> 或 <ol> 标签包裹
 */
class UnwrappedTextFixer extends ErrorFixer {
  canFix() {
    // 只要错误消息中包含 "DIV 元素包含未包裹文本"，就可以修复
    return this.errorMessage.includes("DIV 元素包含未包裹文本");
  }

  fix() {
    // 查找所有 div 元素
    const allDivs = this.document.querySelectorAll("div");
    let fixCount = 0;

    allDivs.forEach((div) => {
      // 检查 div 是否直接包含文本节点（未被标签包裹）
      if (this.hasDirectTextNode(div)) {
        // 确定使用什么标签包裹
        const wrapTag = this.determineWrapTag(div);
        // 包裹文本
        this.wrapTextContent(div, wrapTag);
        fixCount++;
      }
    });

    if (fixCount > 0) {
      this.fixed = true;
      this.fixDescription = `为${fixCount}个DIV元素的文本添加了标签包裹`;
    }

    return this.fixed;
  }

  /**
   * 检查 div 是否直接包含文本节点（未被标签包裹的文本）
   */
  hasDirectTextNode(div) {
    // 遍历 div 的直接子节点
    for (const node of div.childNodes) {
      // 如果是文本节点且包含非空白内容
      if (node.nodeType === 3) {
        // Node.TEXT_NODE
        const textContent = node.textContent.trim();
        if (textContent) {
          return true;
        }
      }
    }
    return false;
  }

  /**
   * 确定使用什么标签包裹文本
   */
  determineWrapTag(div) {
    const className = div.className || "";

    // 根据类名或内容特征决定使用的标签
    // 注意：html2pptx 只支持 <p>、<h1>-<h6>、<ul>、<ol>
    if (
      className.includes("title") &&
      !className.includes("report-title") &&
      !className.includes("page-title")
    ) {
      return "h3";
    } else {
      // 默认使用 p 标签（包括 icon、number、footer 等所有其他情况）
      return "p";
    }
  }

  /**
   * 包裹 div 中的文本内容
   */
  wrapTextContent(div, wrapTag) {
    const newChildren = [];

    // 遍历所有子节点
    div.childNodes.forEach((node) => {
      if (node.nodeType === 3) {
        // 文本节点
        const textContent = node.textContent.trim();
        if (textContent) {
          // 创建包裹元素
          const wrapper = this.document.createElement(wrapTag);
          wrapper.textContent = textContent;
          newChildren.push(wrapper);
        } else if (
          node.textContent.includes("\n") ||
          node.textContent.includes(" ")
        ) {
          // 保留空白节点（用于格式化）
          newChildren.push(node.cloneNode(true));
        }
      } else {
        // 保留元素节点
        newChildren.push(node.cloneNode(true));
      }
    });

    // 清空 div 并添加新的子节点
    div.innerHTML = "";
    newChildren.forEach((child) => {
      div.appendChild(child);
    });
  }
}

/**
 * CSS渐变修复器
 * 处理错误: 将线性渐变和径向渐变转换为单色背景，以及背景图片的处理
 */
class CssGradientFixer extends ErrorFixer {
  canFix() {
    // 检查错误消息是否包含渐变相关错误
    if (this.errorMessage.includes("请使用纯色或边框作为形状")) {
      return true;
    }

    // 主动检测：直接检查HTML内容是否包含渐变或背景图片
    const hasGradient =
      this.htmlContent.includes("linear-gradient") ||
      this.htmlContent.includes("radial-gradient") ||
      this.htmlContent.includes("background-image: url(");

    return hasGradient;
  }

  fix() {
    // 查找所有 <style> 标签
    const styleElements = this.document.querySelectorAll("style");
    let fixCount = 0;

    styleElements.forEach((styleElement) => {
      let cssContent = styleElement.textContent;
      let originalContent = cssContent;

      // 处理线性渐变：linear-gradient()
      console.log(`处理线性渐变`);
      cssContent = this.fixLinearGradients(cssContent);

      // 处理径向渐变：radial-gradient()
      console.log("处理径向渐变");
      cssContent = this.fixRadialGradients(cssContent);

      // 处理背景图片：background-image: url(...)
      cssContent = this.fixBackgroundImages(cssContent);

      // 如果内容有变化，更新样式
      if (cssContent !== originalContent) {
        styleElement.textContent = cssContent;
        fixCount++;
      }
    });

    if (fixCount > 0) {
      this.fixed = true;
      this.fixDescription = `将${fixCount}个样式块中的CSS渐变和背景图片转换为单色背景`;
    }

    return this.fixed;
  }

  /**
   * 修复线性渐变
   * 将 linear-gradient(...) 替换为第一个颜色
   */
  fixLinearGradients(cssContent) {
    // 使用更准确的正则表达式匹配整个background属性值
    const linearGradientRegex = /background:\s*linear-gradient\([^;]+\);/g;

    return cssContent.replace(linearGradientRegex, (match) => {
      // 提取第一个颜色值
      const firstColor = this.extractFirstColor(match);
      return `background: ${firstColor};`;
    });
  }

  /**
   * 修复径向渐变
   * 将 radial-gradient(...) 替换为第一个颜色
   */
  fixRadialGradients(cssContent) {
    // 使用更准确的正则表达式匹配整个background属性值
    const radialGradientRegex = /background:\s*radial-gradient\([^;]+\);/g;

    return cssContent.replace(radialGradientRegex, (match) => {
      // 提取第一个颜色值
      const firstColor = this.extractFirstColor(match);
      return `background: ${firstColor};`;
    });
  }

  /**
   * 从渐变字符串中提取第一个颜色
   */
  extractFirstColor(gradientString) {
    // 匹配十六进制颜色 (#xxx 或 #xxxxxx)
    const hexColorMatch = gradientString.match(/#[0-9a-fA-F]{3,6}/);
    if (hexColorMatch) {
      return hexColorMatch[0];
    }

    // 匹配 rgb/rgba 颜色
    const rgbColorMatch = gradientString.match(/rgba?\([^)]+\)/);
    if (rgbColorMatch) {
      return rgbColorMatch[0];
    }

    // 匹配颜色名称（如 red, blue 等）
    const colorNameMatch = gradientString.match(
      /\b(red|blue|green|yellow|purple|orange|pink|black|white|gray|grey)\b/i
    );
    if (colorNameMatch) {
      return colorNameMatch[0];
    }

    // 如果都没匹配到，返回默认颜色
    return "#000000";
  }

  /**
   * 修复背景图片
   * 将 background-image: url(...) 替换为默认背景色
   */
  fixBackgroundImages(cssContent) {
    // 匹配背景图片：background-image: url(...)
    const backgroundImageRegex = /background-image:\s*url\([^)]*\);/g;

    return cssContent.replace(backgroundImageRegex, () => {
      // 替换为默认的浅灰色背景
      return `background-color: #f5f5f5;`;
    });
  }
}

/**
 * 自动修复HTML文件
 */
async function autoFixHtml(htmlPath, errorMessage, options = {}) {
  const { backup = false } = options;

  // 按优先级尝试各种修复器
  const fixers = [
    new TextElementBorderFixer(errorMessage, htmlPath),
    new UnwrappedTextFixer(errorMessage, htmlPath),
    new CssGradientFixer(errorMessage, htmlPath),
  ];

  let hasAnyFix = false;
  const appliedFixers = [];

  for (const fixer of fixers) {
    console.log(`检查修复器: ${fixer.constructor.name}`);
    if (fixer.canFix()) {
      console.log(`  🎯 使用修复器: ${fixer.constructor.name}`);
      const fixed = fixer.fix();
      if (fixed) {
        // 第一个修复器不备份原文件，后续修复器也不备份
        fixer.save(false);
        hasAnyFix = true;
        appliedFixers.push(fixer.constructor.name);

        // 更新其他修复器的DOM，使它们基于已修复的版本继续工作
        if (appliedFixers.length < fixers.length) {
          const updatedContent = fixer.dom.serialize();
          // 更新后续修复器的DOM
          for (let i = fixers.indexOf(fixer) + 1; i < fixers.length; i++) {
            const nextFixer = fixers[i];
            nextFixer.htmlContent = updatedContent;
            nextFixer.dom = new JSDOM(updatedContent);
            nextFixer.document = nextFixer.dom.window.document;
          }
        }
      }
    }
  }

  if (hasAnyFix) {
    console.log(`  ✅ 已应用修复器: ${appliedFixers.join(", ")}`);
    return true;
  }

  console.log(`  ❌ 未找到适合的修复器`);
  return false;
}

module.exports = {
  autoFixHtml,
  ErrorFixer,
  TextElementBorderFixer,
  UnwrappedTextFixer,
  CssGradientFixer,
};
