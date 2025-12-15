/**
 * HTML to PPTX 转换工具
 *
 * 使用方法：
 * 1. 转换整个文件夹：node convert.js --folder slides --output merged.pptx
 * 2. 转换单个文件：node convert.js --file slide_01_cover.html --output single.pptx
 */

const pptxgen = require("pptxgenjs");
const html2pptx = require("./html2pptx.js");
const { autoFixHtml } = require("./auto_fix.js");
const fs = require("fs");
const path = require("path");

// 尝试加载Playwright，如果没有安装则跳过图表功能
let chromium = null;
try {
  chromium = require("playwright").chromium;
} catch (error) {
  console.log(
    "⚠️  Playwright未安装，图表截取功能不可用。运行: npm install playwright"
  );
}

/**
 * 截取页面中的canvas图表
 */
async function captureCanvasCharts(htmlFile) {
  if (!chromium) {
    return [];
  }

  try {
    const browser = await chromium.launch();
    const page = await browser.newPage();

    const filePath = path.isAbsolute(htmlFile)
      ? htmlFile
      : path.join(process.cwd(), htmlFile);
    await page.goto(`file://${filePath}`);

    // 等待图表渲染完成
    await page.waitForTimeout(1000);

    // 获取所有canvas元素的截图
    const canvasData = await page.evaluate(() => {
      const canvases = Array.from(document.querySelectorAll("canvas"));
      return canvases.map((canvas) => {
        const rect = canvas.getBoundingClientRect();
        return {
          id: canvas.id || canvas.parentElement?.id || "",
          dataUrl: canvas.toDataURL("image/png"),
          position: {
            x: rect.left / 96, // 转换为英寸
            y: rect.top / 96,
            w: rect.width / 96,
            h: rect.height / 96,
          },
        };
      });
    });

    await browser.close();
    return canvasData;
  } catch (error) {
    console.error(`图表截取失败: ${error.message}`);
    return [];
  }
}

/**
 * 从JSON文件加载图表数据
 */
function loadChartsData(chartsFile) {
  if (!chartsFile || !fs.existsSync(chartsFile)) {
    return [];
  }

  try {
    const data = fs.readFileSync(chartsFile, "utf8");
    return JSON.parse(data);
  } catch (error) {
    console.error(`加载图表数据失败: ${error.message}`);
    return [];
  }
}

/**
 * 将图表数据插入到slide中
 */
function insertChartsToSlide(slide, chartsData, placeholders = []) {
  if (!chartsData || chartsData.length === 0) {
    return;
  }

  for (const chart of chartsData) {
    try {
      // 查找匹配的placeholder
      const placeholder = placeholders.find(
        (p) =>
          chart.id &&
          (p.id === chart.id ||
            p.id.includes(chart.id) ||
            chart.id.includes(p.id))
      );

      if (placeholder) {
        // 使用placeholder位置
        slide.addImage({
          data: chart.dataUrl || chart.data_url,
          x: placeholder.x,
          y: placeholder.y,
          w: placeholder.w,
          h: placeholder.h,
        });
      } else {
        // 使用图表自己的位置
        slide.addImage({
          data: chart.dataUrl || chart.data_url,
          x: chart.position.x,
          y: chart.position.y,
          w: chart.position.w,
          h: chart.position.h,
        });
      }
    } catch (error) {
      console.error(`插入图表失败 (${chart.id}): ${error.message}`);
    }
  }
}

/**
 * 尝试转换HTML文件，如果失败则尝试修复后重试
 */
async function tryConvertWithAutoFix(htmlFile, pptx) {
  let lastError = null;

  // 第一次尝试：直接转换
  try {
    const result = await html2pptx(htmlFile, pptx);
    console.log(`  ✓ 直接转换成功`);
    return { success: true, result, method: "direct" };
  } catch (error) {
    lastError = error.message;
    console.log(`  ⚠️  初次转换失败: ${error.message.substring(0, 800)}...`);
  }

  // 第二次尝试：auto_fix修复后重试
  console.log(`  🔧 尝试auto_fix修复...`);
  try {
    const fixed = await autoFixHtml(htmlFile, lastError, {
      backup: true,
    });

    if (fixed) {
      console.log(`  ✓ auto_fix修复成功，重新转换...`);
      try {
        const result = await html2pptx(htmlFile, pptx);
        console.log(`  ✓ 修复后转换成功`);
        return { success: true, result, method: "auto_fix" };
      } catch (retryError) {
        lastError = retryError.message;
        console.log(
          `  ✗ 修复后仍转换失败: ${retryError.message.substring(0, 80)}...`
        );
      }
    } else {
      console.log(`  ⚠️  auto_fix无法修复此错误`);
    }
  } catch (fixError) {
    console.log(`  ✗ auto_fix修复过程出错: ${fixError.message}`);
  }

  // 修复失败，返回最后的错误
  return {
    success: false,
    error: lastError,
    method: "failed",
  };
}

/**
 * 转换单个HTML文件为PPTX
 */
async function convertSingleFile(htmlFile, outputFile) {
  console.log(`\n开始转换文件: ${htmlFile}`);

  const pptx = new pptxgen();

  // 定义1280x720布局
  pptx.defineLayout({
    name: "CUSTOM_1280x720",
    width: 13.33, // 1280px ÷ 96 = 13.33英寸
    height: 7.5, // 720px ÷ 96 = 7.5英寸
  });
  // 定义1600x900布局
  pptx.defineLayout({
    name: "CUSTOM_1600x900",
    width: 16.67, // 1600px ÷ 96 = 16.67英寸
    height: 9.38, // 900px ÷ 96 = 9.38英寸
  });

  pptx.layout = "CUSTOM_1600x900";

  // 尝试转换（包含auto_fix）
  const convertResult = await tryConvertWithAutoFix(htmlFile, pptx);

  if (!convertResult.success) {
    console.error(`✗ 转换失败: ${convertResult.error}`);
    return false;
  }

  const {
    result: { slide, placeholders },
    method,
  } = convertResult;
  console.log(`✓ 成功转换 (方法: ${method}): ${htmlFile}`);

  let chartsData = [];

  // 默认尝试加载对应的图表JSON文件
  const chartsFile = htmlFile.replace(".html", ".charts.json");
  chartsData = loadChartsData(chartsFile);

  if (chartsData.length > 0) {
    console.log(`  从文件加载 ${chartsData.length} 个图表`);
  }
  // 如果没有JSON文件且支持Playwright，尝试实时截取
  else if (chromium) {
    console.log(`  正在截取图表...`);
    chartsData = await captureCanvasCharts(htmlFile);
    if (chartsData.length > 0) {
      console.log(`  截取到 ${chartsData.length} 个图表`);
    }
  }

  // 插入图表
  if (chartsData.length > 0) {
    insertChartsToSlide(slide, chartsData, placeholders);
  }

  await pptx.writeFile({ fileName: outputFile });
  console.log(`\n✓ PPTX 文件已保存: ${outputFile}`);

  return true;
}

/**
 * 转换文件夹中所有HTML文件为一个PPTX
 */
async function convertFolder(folderPath, outputFile) {
  const htmlFiles = fs
    .readdirSync(folderPath)
    .filter(
      (file) =>
        file.endsWith(".html") &&
        !file.endsWith(".backup") &&
        !file.startsWith("_skip_")
    )
    .sort()
    .map((file) => path.join(folderPath, file));

  if (htmlFiles.length === 0) {
    console.error(`✗ 文件夹 "${folderPath}" 中没有找到HTML文件`);
    process.exit(1);
  }

  console.log(`\n找到 ${htmlFiles.length} 个HTML文件:`);
  htmlFiles.forEach((file, index) => {
    console.log(`  ${index + 1}. ${path.basename(file)}`);
  });

  const pptx = new pptxgen();

  // 定义1280x720布局
  pptx.defineLayout({
    name: "CUSTOM_1280x720",
    width: 13.33, // 1280px ÷ 96 = 13.33英寸
    height: 7.5, // 720px ÷ 96 = 7.5英寸
  });
  pptx.defineLayout({
    name: "CUSTOM_1600x900",
    width: 16.67, // 1600px ÷ 96 = 16.67英寸
    height: 9.38, // 900px ÷ 96 = 9.38英寸
  });

  pptx.layout = "CUSTOM_1600x900";

  console.log("\n开始转换...\n");

  const results = {
    success: 0,
    failed: 0,
    direct: 0, // 直接转换成功
    autoFixed: 0, // auto_fix修复后成功
    failedFiles: [],
  };

  for (let i = 0; i < htmlFiles.length; i++) {
    const htmlFile = htmlFiles[i];
    const fileName = path.basename(htmlFile);

    console.log(`[${i + 1}/${htmlFiles.length}] ${fileName}`);

    // 尝试转换（包含auto_fix）
    const convertResult = await tryConvertWithAutoFix(htmlFile, pptx);

    if (convertResult.success) {
      const {
        result: { slide, placeholders },
        method,
      } = convertResult;
      results.success++;

      if (method === "direct") {
        results.direct++;
      } else if (method === "auto_fix") {
        results.autoFixed++;
      }

      let chartsData = [];

      // 默认尝试加载对应的图表JSON文件
      const chartsFile = htmlFile.replace(".html", ".charts.json");
      chartsData = loadChartsData(chartsFile);

      if (chartsData.length > 0) {
        console.log(`    从文件加载 ${chartsData.length} 个图表`);
      }
      // 如果没有JSON文件且支持Playwright，尝试实时截取
      else if (chromium) {
        console.log(`    正在截取图表...`);
        chartsData = await captureCanvasCharts(htmlFile);
        if (chartsData.length > 0) {
          console.log(`    截取到 ${chartsData.length} 个图表`);
        }
      }

      // 插入图表
      if (chartsData.length > 0) {
        insertChartsToSlide(slide, chartsData, placeholders);
      }
    } else {
      console.error(`  ✗ 最终转换失败`);
      results.failed++;
      results.failedFiles.push({
        name: fileName,
        error: convertResult.error,
      });

      console.log(`    ⏭️  跳过此文件，继续处理下一个...`);
    }

    console.log("");
  }

  // 只有成功转换至少一个文件才生成PPTX
  if (results.success > 0) {
    await pptx.writeFile({ fileName: outputFile });
    console.log(`\n✓ PPTX文件已保存: ${outputFile}`);
    console.log(`  包含幻灯片: ${results.success} 张`);
    console.log(`  直接转换成功: ${results.direct} 个`);
    console.log(`  auto_fix修复后成功: ${results.autoFixed} 个`);

    if (results.failed > 0) {
      console.log(`  跳过文件: ${results.failed} 个`);
      console.log(`\n⚠️  跳过的文件详情:`);
      results.failedFiles.forEach((file) => {
        console.log(`    • ${file.name}: ${file.error}`);
      });
    }
  } else {
    console.error(`\n✗ 所有文件转换失败，无法生成PPTX`);
    console.log(`\n失败文件详情:`);
    results.failedFiles.forEach((file) => {
      console.log(`  • ${file.name}: ${file.error}`);
    });
    return false;
  }

  return results.success > 0;
}

/**
 * 解析命令行参数
 */
function parseArgs() {
  const args = process.argv.slice(2);
  const options = {
    mode: null,
    input: null,
    output: "output.pptx",
  };

  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--folder" && args[i + 1]) {
      options.mode = "folder";
      options.input = args[i + 1];
      i++;
    } else if (args[i] === "--file" && args[i + 1]) {
      options.mode = "file";
      options.input = args[i + 1];
      i++;
    } else if (args[i] === "--output" && args[i + 1]) {
      options.output = args[i + 1];
      i++;
    }
  }

  return options;
}

/**
 * 主函数
 */
async function main() {
  const options = parseArgs();

  // 显示帮助信息
  if (!options.mode) {
    console.log(`
HTML to PPTX 转换工具 (自动图表截取 + 智能修复)

使用方法:
  转换整个文件夹:
    node convert.js --folder <文件夹路径> --output <输出文件.pptx>

  转换单个文件:
    node convert.js --file <HTML文件路径> --output <输出文件.pptx>

示例:
  node convert.js --folder slides --output merged.pptx
  node convert.js --file slides/slide_01_cover.html --output single.pptx

参数:
  --folder         指定包含HTML文件的文件夹路径
  --file           指定单个HTML文件路径
  --output         指定输出的PPTX文件名（可选，默认为 output.pptx）

转换策略:
  ✓ 首先尝试直接转换
  ✓ 转换失败时自动调用auto_fix修复
  ✓ 修复后重新尝试转换
  ✓ 无法修复的文件自动跳过
  ✓ 显示详细的转换统计信息

图表功能:
  ✓ 默认启用图表截取功能
  ✓ 自动检测Canvas图表元素
  ✓ 优先使用.charts.json文件，否则实时截取
  ✓ 支持占位符匹配
  ✓ 向后兼容，无Playwright时跳过图表
    `);
    process.exit(0);
  }

  // 验证输入路径
  if (!fs.existsSync(options.input)) {
    console.error(`✗ 路径不存在: ${options.input}`);
    process.exit(1);
  }

  // 执行转换
  let success = false;

  if (options.mode === "folder") {
    success = await convertFolder(options.input, options.output);
  } else if (options.mode === "file") {
    success = await convertSingleFile(options.input, options.output);
  }

  process.exit(success ? 0 : 1);
}

// 运行
main().catch((error) => {
  console.error("\n✗ 发生错误:", error.message);
  process.exit(1);
});
