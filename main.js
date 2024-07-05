const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');


console.log("path.resolve(__dirname);", path.resolve(__dirname))

const sourceDir = './docs';
const targetDir = './md/test';
const targetJSONDir = './md/JSON';
const targetDirH = './html/test';


// 确保目标目录存在
function exsureDirExists(dir) {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
}

exsureDirExists(targetDir);
exsureDirExists(targetDirH);
exsureDirExists(targetJSONDir);

// 读取目录中的所有.docx文件
fs.readdir(sourceDir, (err, files) => {
    if (err) {
        console.error("Could not list the directory.", err);
        process.exit(1);
    }
    console.log("file", files)
    files.forEach((file) => {
        var arrJSON = [];
        if (path.extname(file) === '.docx') {
            const sourceFile = path.join(sourceDir, file);
            const targetMdDir = path.join(targetDir, path.basename(file, '.docx') + '.md');
            const targetHTMLDir = path.join(targetDirH, path.basename(file, '.docx') + '.html');
            const targetTextDir = path.join(targetJSONDir, path.basename(file, '.docx') + '.json');

            // 使用mammoth读取.docx文件并转换为HTML
            // mammoth.convertToHtml({ path: sourceFile })
            //     .then((result) => {
            //         const html = result.value; // 转换后HTML内容
            //         fs.writeFile(targetHTMLDir, html, (err) => {
            //             if (err) throw err;
            //             console.log(`[HTML]: The file ${file} has been saved as ${path.basename(targetHTMLDir)}`);
            //         });
            //     })
            //     .done();

            // 使用mammoth读取.docx文件并转换为MD
            mammoth.convertToMarkdown({ path: sourceFile })
                .then((result) => {
                    var mdContent = result.value; // 转换后内容
                    // mdContent = mdContent.replace(/(\d+(\\\.\d+)*)\.? __([^_]+)__/g, function(match, p1, p2, p3) {
                    //     var count = (p1.match(/\./g) || []).length + 1; // // <a>标签"."的数量
                    //     return '#'.repeat(count + 1) + ' ' + p3.trim(); // 根据层级生成相应数量的#
                    // });
                    
                    mdContent = mdContent.replace(/(<a[^>]*><\/a>)?(\d+(?:\\\.\d+)*)\\?\.? __([^_]+)__/g, function (match, p1, p2, str3, p4) {
                        var p3 = encodeURI(str3);
                        // var p3 = str3;
                        var count = (p2.match(/\\\./g) || []).length + 1;
                        // var heading = p1 ? p1.replace(/heading_\d+/, `heading_${encodeURI(p3)}`) + '\n' : '';
                        var heading = p1 ? p1.replace(/heading_\d+/, `heading_${p3}`) + '\n' : '';
                        var content = '#'.repeat(count + 1) + ' ' + str3.trim(); 
                        arrJSON.push({
                            key: `key_${p3}`,
                            href: `#heading_${p3}`,
                            title: p3.trim()
                        })
                        return heading + content;
                    });

                    // console.log(mdContent);
                    fs.writeFile(targetMdDir, mdContent, (err) => {
                        if (err) throw err;
                        console.log(`[MD]: The file ${file} has been saved as ${path.basename(targetMdDir)}`);
                    });
                    fs.writeFile(targetTextDir, JSON.stringify(arrJSON), (err) => {
                        if (err) throw err;
                        console.log(`[text]: The file ${file} has been saved as ${path}`)
                    });
                })
                .done();
        }
    });
});


// async function convertWordToMd(wordFilePath, outputMdPath) {
//     console.log("拿到没？？？", wordFilePath)
//   try {
//     const result = await mammoth.convertToMarkdown({path: wordFilePath});
//     const mdContent = result.value; //
//     // console.log("这就是转换后的Markdown内容", mdContent)
//     fs.writeFileSync(outputMdPath, mdContent);
//     console.log(`转换完成！Markdown文件已保存至：${outputMdPath}`);
//   } catch (error) {
//     console.error('转换过程中发生错误：', error);
//   }
// }
// const wordFilePath = `./document.docx`;
// const outputMdPath = `./document.md`;
// convertWordToMd(wordFilePath, outputMdPath);

// npm install mammoth