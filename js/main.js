"use strict";
async function changeEventHandler(e) {

    async function writeStringToFile(path, data) {
        await pyodide.runPythonAsync(
            `
        import os,base64
        fpath = '${path}'
        _path = os.path.split('${path}')[0]
        if not os.path.exists(_path):
            os.makedirs(_path)
        with open(fpath,'wb') as f:
            f.write(base64.b64decode("${data}"))
        `
        )
        await pyodide.runPythonAsync(
            `
        import os
        print(os.listdir(os.path.split('${path}')[0]))
        `
        )
    }
    console.log(e.target.name);
    let _prompt = ""
    for (let i = 0; i < e.target.files.length; i++) {
        let f = e.target.files[i]
        console.log(f.webkitRelativePath);
        _prompt += "\n"+f.webkitRelativePath
        let reader = new FileReader();
        reader.readAsBinaryString(f)
        reader.onload = async _e => {
            let data = btoa(reader.result);
            console.log(f.webkitRelativePath, data.length)
            //let stream = pyodide.FS.open(f.webkitRelativePath,'w+')
            //pyodide.FS.write(stream,data,0,data.length,0)
            //pyodide.FS.close(stream)
            //pyodide.FS.writeFile("www.txt",data)
            await writeStringToFile(f.webkitRelativePath, data)

        }
    }
    console.log('paper gen...')
    let fname = prompt("从以下路径中选择目标md文件完整path输入:"+_prompt)
    await execPaperGen(e.target.name, fname)
}
$('input[type="file"]').change(changeEventHandler)


function sleep(s) {
    return new Promise((resolve) => setTimeout(resolve, s));
}


async function init() {
    globalThis.pyodide = await loadPyodide({
        indexURL: "https://cdn.jsdelivr.net/pyodide/v0.19.0/full/",
    });
    namespace = pyodide.globals.get("dict")();
}
async function loadDeps() {
    await pyodide.loadPackage(["micropip"])
    await pyodide.runPythonAsync(
        `
        import micropip
        libs = [
            'bibtexparser-1.2.0-py3-none-any',
            'python_docx-0.8.11-py3-none-any',
            'md2paper-1.0.0-py3-none-any'
        ]
        
        async def mi(lib:str):
            await micropip.install(f'./libs/{lib}.whl')
            print(lib)
            return
        await mi(libs[0])
        await mi(libs[1])
        await mi(libs[2])
        `,
        namespace
    ).then(() => {
        console.log('deps are ready')
    });
}
async function execPaperGen(ftype, fname) {
    $("#msg").text('正在生成……')
    if (ftype !== "grad" && ftype !== "trans") {
        throw new Error("invalid ftype" + ftype)
    }
    const mapping = {
        "grad": {
            "tpl": "毕业设计（论文）模板-docx.docx",
            "class": "GraduationPaper"
        },
        "trans": {
            "tpl": "外文翻译模板-docx.docx",
            "class": "TranslationPaper"
        }
    }
    await pyodide.runPythonAsync(
        `
            from io import StringIO,BytesIO
            from md2paper import GraduationPaper,TranslationPaper
            from pyodide.http import pyfetch
            print('md2paper ok')
            r1 = await pyfetch("./word-template/${mapping[ftype]['tpl']}")
            fname = "${fname}"

            p = ${mapping[ftype]['class']}()
            p.load_md(fname)
            p.load_contents()
            p.compile()
            tpl = BytesIO(await r1.bytes())
            p.render(tpl,"out.docx")
            print('render ok')
            `,
        namespace
    )
    const data = pyodide.FS.readFile("out.docx")
    const blob = new Blob([data.buffer], {
        type: 'application/msword;charset=utf-8'
    });
    const fileName = `out.docx`;
    const link = document.createElement('a');
    link.href = window.URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
    window.URL.revokeObjectURL(link.href);
    $("#msg").text('')
}
const start = async function () {
    $("#msg").text('等待依赖加载……时长主要取决于网络状况')
    await init();
    await loadDeps();
    $("#msg").text('依赖加载完成')
    $('input[type="file"]').removeAttr("disabled");
    setTimeout(() => {
        $("#msg").text('')
    }, 5000);
}
/*
- 装依赖
- 用户点击选择文件，选择md文件/包含图片的整个文件夹
- 生成，造出docx并下载
*/
let namespace;
// Call start
start();