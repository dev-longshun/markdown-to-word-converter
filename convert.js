const fs = require("fs").promises;
const path = require("path");
const { convertMarkdownToDocx } = require("@mohtasham/md-to-docx");

// --- 配置区域 ---
const SOURCE_DIR = "markdown_files";
const OUTPUT_DIR = "word_documents";

const styleOptions = {
  styles: {
    // ------------------- 基础样式 -------------------
    normal: {
      font: "华文宋体", // STSong
      size: 28, // 四号 (14pt * 2)
      paragraph: {
        alignment: "both",
        spacing: { line: 300 }, // 1.25倍行距
        indent: { firstLine: 840 }, // 首行缩进2字符
      },
    },

    // ------------------- 标题样式 -------------------
    heading_1: {
      font: "华文黑体", // STHeiti
      size: 32, // 三号 (16pt * 2)
      bold: true,
      paragraph: {
        alignment: "center",
        spacing: { before: 320, after: 160 }, // 增加段前/段后间距
      },
    },
    heading_2: {
      font: "华文宋体", // STSong
      size: 30, // 小三 (15pt * 2)
      bold: true,
      paragraph: {
        alignment: "start",
        spacing: { before: 240, after: 120 },
      },
    },
    heading_3: {
      font: "华文宋体", // STSong
      size: 28, // 四号 (14pt * 2)
      bold: true,
      paragraph: {
        alignment: "both",
        spacing: { before: 200, after: 100 },
      },
    },

    // ------------------- 新增与优化样式 -------------------

    // 【优化】用于英文/数字的样式 (通过 `code` 语法触发)
    inline_code: {
      font: "Times New Roman", // 指定英文字体
      size: 28, // 与正文四号字 (14pt) 保持一致
    },

    // 【新增】用于图表标题的样式 (通过 *斜体* 语法触发)
    emphasis: {
      font: "华文宋体", // STSong
      size: 21, // 五号 (10.5pt * 2)
      italic: false, // 我们只是借用斜体语法，实际并不需要斜体效果
      paragraph: {
        alignment: "center",
        indent: { firstLine: 0 }, // 标题不缩进
      },
    },

    // 【新增】表格内文字样式
    table_cell: {
      font: "华文宋体",
      size: 21, // 五号 (10.5pt * 2)
      paragraph: {
        alignment: "center", // 表格内容居中
        indent: { firstLine: 0 }, // 不缩进
      },
    },
  },
};

async function processFiles() {
  try {
    await fs.mkdir(SOURCE_DIR, { recursive: true });
    await fs.mkdir(OUTPUT_DIR, { recursive: true });
    console.log(`扫描源文件夹: ./${SOURCE_DIR}`);
    console.log(`检查输出文件夹: ./${OUTPUT_DIR}`);

    const allFilesInSource = await fs.readdir(SOURCE_DIR);
    const markdownFiles = allFilesInSource.filter(
      (file) => path.extname(file).toLowerCase() === ".md"
    );

    if (markdownFiles.length === 0) {
      console.log("在源文件夹中没有找到 Markdown 文件。");
      return;
    }
    console.log(`发现 ${markdownFiles.length} 个 Markdown 文件。开始处理...`);
    console.log("---");

    for (const mdFile of markdownFiles) {
      const sourceFilePath = path.join(SOURCE_DIR, mdFile);
      const baseName = path.basename(mdFile, ".md");
      const targetDocxFile = `${baseName}.docx`;
      const outputFilePath = path.join(OUTPUT_DIR, targetDocxFile);

      try {
        await fs.access(outputFilePath);
        console.log(`[跳过] '${targetDocxFile}' 已经存在。`);
      } catch (error) {
        if (error.code === "ENOENT") {
          console.log(`[转换中] 正在转换 '${mdFile}'...`);
          try {
            const markdownContent = await fs.readFile(sourceFilePath, "utf-8");
            const docxBlob = await convertMarkdownToDocx(
              markdownContent,
              styleOptions
            );
            const docxBuffer = Buffer.from(await docxBlob.arrayBuffer());
            await fs.writeFile(outputFilePath, docxBuffer);
            console.log(`[成功] 已创建 '${targetDocxFile}'。`);
          } catch (conversionError) {
            console.error(
              `[失败] 转换 '${mdFile}' 时发生错误:`,
              conversionError
            );
          }
        } else {
          console.error(`检查文件 '${targetDocxFile}' 时发生未知错误:`, error);
        }
      }
    }
    console.log("---");
    console.log("所有文件处理完毕！");
  } catch (err) {
    console.error("脚本执行过程中发生严重错误:", err);
  }
}

processFiles();
