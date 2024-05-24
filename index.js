const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const readline = require('readline');

// https://stackoverflow.com/questions/68604699/node-js-pkg-executable-not-able-to-read-files-from-outside
// const directoryPath = __dirname;
/**
 * `__dirname` 在普通的 Node.js 环境中指向当前脚本所在的目录路径。
 * 但在打包后的可执行文件中，这个路径指向的是虚拟文件系统中的临时路径，而不是你实际的项目目录。
 * `process.cwd()` 返回的是当前工作目录的路径，这个路径是在运行可执行文件时用户所在的目录。
 * 因此，使用 `process.cwd()` 可以确保路径解析正确，指向实际的项目目录，而不是临时目录。
 */
const directoryPath = process.cwd();

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

const askQuestion = (query) => new Promise((resolve) => rl.question(query, resolve));

const isHiddenFile = (filePath) => {
  return new Promise((resolve, reject) => {
    fs.stat(filePath, (err, stats) => {
      if (err) {
        return reject(err);
      }

      // Check if the file is hidden
      if (process.platform === 'win32') {
        // Windows: Use file attributes to check if the file is hidden
        const isHidden = (stats.mode & 0o1000000) !== 0;
        resolve(isHidden);
      } else {
        // Unix-like: Hidden files start with a dot
        const isHidden = path.basename(filePath).startsWith('.');
        resolve(isHidden);
      }
    });
  });
};

const waitForKeyPress = async () => {
  console.log('\n按任意键退出……');
  process.stdin.setRawMode(true);
  return new Promise((resolve) =>
    process.stdin.once('data', () => {
      process.stdin.setRawMode(false);
      resolve();
    }),
  );
};

// 新增递归读取目录的函数
const readDirRecursive = async (dir, fileList = []) => {
  const files = await fs.promises.readdir(dir, { withFileTypes: true });
  for (const file of files) {
    const fullPath = path.join(dir, file.name);
    if (file.isDirectory()) {
      await readDirRecursive(fullPath, fileList);
    } else {
      fileList.push(fullPath);
    }
  }
  return fileList;
};

(async () => {
  try {
    console.log(`
    [-------------------○-------------------]
    欢迎使用 Excel文件切片工具 by 金钊
    [=================<说明>=================]
    起始行是键值，剩余行是元素。空行会被忽略。
    文件名可以是相对路径，比如../test/myData。
    留空则遍历处理当前目录及子目录下所有excel文件。
    [=================<注意>=================]
    安全性：不会影响隐藏文件，包括excel临时文件。
    [-------------------○-------------------]
    `);

    const fileName = await askQuestion('请输入文件名（不包括拓展名，可以留空）: ');

    // 指定了文件名的情况
    let filePath;
    if (fileName) {
      filePath = path.join(directoryPath, `${fileName}.xlsx`);
      if (!fs.existsSync(filePath)) {
        filePath = path.join(directoryPath, `${fileName}.xls`);
      }
      if (!fs.existsSync(filePath)) {
        console.error(`文件不存在，没有${fileName}.xlsx或${fileName}.xls文件。`);
        await waitForKeyPress(); // Wait for key press before exiting
        process.exit(1);
      }
    }

    const startInput = await askQuestion('请输入起始行（从1开始）: ');
    const endInput = await askQuestion('请输入结束行（包含该行，可以留空）: ');

    const start = parseInt(startInput, 10);
    const end = endInput ? parseInt(endInput, 10) : undefined;

    if (
      isNaN(start) ||
      start < 1 ||
      (end !== undefined && isNaN(end)) ||
      (end !== undefined && end < start)
    ) {
      console.error('请提供正确的起始和结尾行序号。');
      await waitForKeyPress(); // Wait for key press before exiting
      process.exit(1);
    }

    /**
     * @param {string} file 文件名
     * @param {xlsx.WorkBook} workbook 工作簿
     * @param {number} start 起始行（从1开始）
     * @param {number} end 结束行（可选）
     */
    const processSheets = (file, workbook, start, end) => {
      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        // 将工作表转换为JSON，默认会读取所有行
        // header: 1 - 将每一行的数据作为一个数组返回。
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
        // 选择需要的行
        const filteredData = jsonData.slice(start - 1, end);
        if (filteredData.length === 0) {
          console.warn(`文件 ${file} 没有足够的行来进行切片 ${start}行 到 ${end}行。`);
          return;
        }
        // 将数据转换为JSON格式
        const headers = filteredData[0]; // 获取表头
        const data = filteredData.slice(1); // 获取数据部分

        const outputData = data
          .map((row) => {
            let element = {};
            // 将每一行的数据与表头对应
            row.forEach((cell, index) => {
              const key = headers[index];
              element[key] = cell;
            });
            return element;
          })
          .filter((row) => Object.keys(row).length > 0); // 过滤空对象;

        // 输出JSON文件
        const outputFilePath = file.replace(path.extname(file), '.json');
        const outputFileName = path.relative(directoryPath, outputFilePath);
        fs.writeFileSync(outputFilePath, JSON.stringify(outputData, null, 2));
        console.log(`已将 ${file} 转换为 ${outputFileName}。`);
      });
    };

    console.log('开始处理文件……');
    console.log(fileName);
    // 指定了文件名的情况
    if (fileName) {
      const workbook = xlsx.readFile(filePath);
      processSheets(`${fileName}.xlsx`, workbook, start, end);
      await waitForKeyPress(); // Wait for key press before exiting
      rl.close();
      return;
    }

    // 没有指定文件名，读取目录及子目录中的所有文件
    const allFiles = await readDirRecursive(directoryPath);
    // 检测文件是否有效（Excel文件，不是临时文件，不是隐藏文件）
    const checks = allFiles.map(async (file) => {
      const isExcelFile = file.endsWith('.xlsx') || file.endsWith('.xls');
      const isNotTempFile = !file.startsWith('~$');
      const isNotHiddenFile = !(await isHiddenFile(file));
      return {
        file,
        isValid: isExcelFile && isNotTempFile && isNotHiddenFile,
      };
    });

    // 等待所有检查完成并过滤有效文件
    const results = await Promise.all(checks);
    const excelFiles = results.filter((result) => result.isValid).map((result) => result.file);

    // 处理每个Excel文件
    excelFiles.forEach((file) => {
      const workbook = xlsx.readFile(file);
      processSheets(file, workbook, start, end);
    });
    await waitForKeyPress(); // Wait for key press before exiting
  } catch (error) {
    console.error('发生错误:', error);
    await waitForKeyPress(); // Wait for key press before exiting
  } finally {
    rl.close();
  }
})();
