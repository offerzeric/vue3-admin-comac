<script lang="ts" setup>
import XLSX from "xlsx"
import axios from "axios"
import { onMounted } from "vue"
import { ElMessage, ElLoading } from "element-plus"

onMounted(() => {
  let result_all_xlsxs_names = []
  /**
   * 索引需要添加交互的html元素
   */
  const sourceFilesUploader = document.querySelector("#source-files-uploader") as any
  const resultFilesUploader = document.querySelector("#result-files-uploader") as any
  const startCodeButton = document.querySelector(".start_code_button") as any
  // var downloadCodeButton = document.querySelector('.download_code_button');
  /**
   * 文件上传和源文件列表显示
   */
  sourceFilesUploader?.addEventListener("change", () => {
    //拿到上传在浏览器缓存区的文件
    const curFiles = sourceFilesUploader.files
    sourceFilesUploader.files = curFiles
    console.log(curFiles)
    const sourceListGroup = document.querySelector(".source-list-group")
    const container = document.getElementById("source-excel-panel")
    const btnsPanel = document.getElementById("source-excel-btns")
    if ((curFiles as any).length === 0) {
      console.log("没有选择文件上传，删除已经在源文件列表中的上传历史.")
      //清空excel预览面板的显示
      // eslint-disable-next-line prettier/prettier
      sourceFilesUploader.files = curFiles
      // eslint-disable-next-line prettier/prettier
      ;(container as HTMLElement).innerHTML = ""
      // eslint-disable-next-line prettier/prettier
      ;(btnsPanel as HTMLElement).innerHTML = ""
      // eslint-disable-next-line prettier/prettier
      ;(sourceListGroup as HTMLElement).innerHTML =
        '<a disabled class="source-list-group-start custom-page-title"' + ">源文件列表</a>"
    } else {
      console.log("源文件上传完成后，展示源文件列表.")
      //调用接口进行上传和预览表格
      excel_and_list_and_upload(sourceListGroup, curFiles)
    }
  })

  /**
   * 网页端预览编码excel表格
   */
  resultFilesUploader?.addEventListener("change", () => {
    //拿到上传在浏览器缓存区的文件
    const curFiles = resultFilesUploader.files
    resultFilesUploader.files = curFiles
    console.log(curFiles)
    const resultListGroup = document.querySelector(".result-list-group") as any
    const container = document.getElementById("result-excel-panel") as any
    const btnsPanel = document.getElementById("result-excel-btns") as any
    if (curFiles.length === 0) {
      console.log("没有选择文件上传，删除已经在结果文件列表中的上传历史.")
      //清空excel预览面板的显示
      resultListGroup.files = null
      container.innerHTML = ""
      btnsPanel.innerHTML = ""
      resultListGroup.innerHTML = '<a disabled class="result-list-group-start custom-page-title"' + ">结果文件列表</a>"
    } else {
      console.log("结果文件上传完成后，展示源文件列表.")
      excel_and_list_and_preview(resultListGroup, curFiles)
    }
  })

  /**
   * 给按钮绑定编码接口
   */
  startCodeButton?.addEventListener("click", () => {
    startCode()
  })

  async function excel_and_list_and_upload(sourceListGroup: any, curFiles: any) {
    //首先恢复初始样式
    sourceListGroup.innerHTML = '<a disabled class="source-list-group-start custom-page-title"' + ">源文件列表</a>"
    for (let i = 0; i < curFiles.length; i++) {
      //文件添加到源文件预览
      const temp = print_excel_panel_and_list(curFiles[i], "source-excel-panel", "source-excel-btns")
      sourceListGroup.appendChild(document.createElement("br"))
      sourceListGroup.appendChild(temp)
    }

    //源文件上传到后端
    console.log("源文件开始上传至服务器.")
    await callUploadFiles()
  }

  /**
   * 用户上传任意表格文件后进行excel预览
   * @param {*} sourceListGroup
   * @param {*} curFiles
   * @returns source list group view
   */
  async function excel_and_list_and_preview(resultListGroup: any, curFiles: any) {
    //首先恢复初始样式
    resultListGroup.innerHTML = '<a disabled class="result-list-group-start custom-page-title"' + ">结果文件列表</a>"

    for (let i = 0; i < curFiles.length; i++) {
      //文件添加到结果文件预览
      const temp = print_excel_panel_and_list(curFiles[i], "result-excel-panel", "result-excel-btns")
      resultListGroup.appendChild(document.createElement("br"))
      resultListGroup.appendChild(temp)
    }
  }

  /**
   * 在源文件列表列出并预览excel文件
   * @param {*} fileTemp
   */
  function print_excel_panel_and_list(fileTemp: any, excelPanel: any, excelBtns: any) {
    //创建a标签用于文件列表展示
    const temp = document.createElement("button") as HTMLButtonElement
    //a标签的样式
    //a标签的id
    temp.id = fileTemp.name
    temp.innerHTML = fileTemp.name
    temp.style.fontSize = "14px"
    temp.className = "el-button"
    //转换成可预览的excel
    temp.addEventListener("click", async (e: any) => {
      e.preventDefault()
      //excel预览框
      const container = document.getElementById(excelPanel) as any
      //sheet 按钮栏
      const btns = document.getElementById(excelBtns) as any
      //每次添加前首先清空上一次已经添加的文件
      container.innerHTML = ""
      btns.innerHTML = ""
      //将文件转成二进制数组方便xlsx插件转换
      const data = await fileTemp.arrayBuffer()
      //workbook为sheet页面
      const workbook = XLSX.read(data)
      const wsnames = workbook.SheetNames
      console.log(wsnames)
      // 对所有sheet进行按钮添加
      for (const sheetName of wsnames) {
        console.log(sheetName)
        const btn = document.createElement("button")
        //按钮的样式
        btn.className = "custom_button_sheet el-button"
        btn.style.fontSize = "14px"
        btn.innerHTML = sheetName
        //给这个sheet按钮添加点击事件
        btn.addEventListener("click", () => {
          container.innerHTML = ""
          const worksheet = workbook.Sheets[sheetName]
          const div_temp = document.createElement("div")
          div_temp.style.fontSize = "14px"
          //调用sheet转html
          div_temp.innerHTML = XLSX.utils.sheet_to_html(worksheet, {
            header: "Sheet Name: " + sheetName
          })
          container.appendChild(div_temp)
        })
        btns.appendChild(btn)
      }
    })
    return temp
  }

  /**
   * 上传文件到服务器
   * @param {*} curFiles
   */
  async function callUploadFiles() {
    const domain = "http://192.168.3.9:9092"
    // const domain = "http://127.0.0.1:9092"
    console.log(domain + "/code/do424Upload")
    const customForm = document.querySelector(".custom_form") as any
    customForm.addEventListener("submit", () => {
      const formData = new FormData(customForm)
      const loadingInstance = ElLoading.service({ fullscreen: true })
      axios
        .post(domain + "/code/do424Upload", formData, {
          headers: {
            "Content-Type": "multipart/form-data",
            "Access-Control-Allow-Origin": "*",
            "Cache-Control": "no-cache"
          },
          validateStatus: function (status) {
            return (status >= 200 && status < 300) || status == 304
          }
        })
        .then((res) => {
          loadingInstance.close()
          console.log(res)
          if (res.data.flag) {
            ElMessage.success("源文件上传成功")
          } else {
            ElMessage.error("源文件上传失败，请稍后再试")
          }
          console.log("上传程序结束")
        })
        .catch(() => {
          ElMessage.error("源文件上传失败，请稍后再试")
          console.log("上传程序结束")
        })
    })
  }

  /**
   * 开始424编码
   */
  async function startCode() {
    console.log("开始对源文件编码.")
    const domain = "http://192.168.3.9:9092"
    // const domain = "http://127.0.0.1:9092"
    console.log(domain + "/code/do424Code")
    const loadingInstance = ElLoading.service({ fullscreen: true })

    //调用后端424编码接口
    await axios
      .get(domain + "/code/do424Code", {
        headers: {
          "Access-Control-Allow-Origin": "*",
          "Cache-Control": "no-cache"
        },
        //验证后端返回状态
        validateStatus: function (status) {
          return (status >= 200 && status < 300) || status == 304
        }
      })
      .then((res) => {
        loadingInstance.close()
        //res即接收到的结果
        console.log(res)
        if (res.data.flag == 1) {
          //后端编码成功，显示debug信息，下载结果文件
          ElMessage.success("编码成功")
          result_all_xlsxs_names = res.data.result_all_xlsxs
          showCodeDebugMsg(res.data.result_all_xlsxs)
          showResultCodeList(result_all_xlsxs_names)
        } else if (res.data.flag == 2) {
          //后端编码失败，流程中出现错误，显示目前已经产生的debug信息
          ElMessage.error("编码失败")
          result_all_xlsxs_names = res.data.result_all_xlsxs
          showCodeDebugMsg(res.data.result_all_xlsxs)
        } else {
          ElMessage.error("编码失败，请稍后再试")
        }
        console.log("424编码程序结束.")
      })
      .catch((error) => {
        console.error(error)
        ElMessage.error("编码失败，请稍后再试")
        console.log("424编码程序结束.")
      })
  }

  /**
   * 下载编码后的文件并将该文件添加到结果文件列表
   * @param {*} result_all_xlsxs_names
   */
  function showResultCodeList(result_all_xlsxs_names: any) {
    for (const each of result_all_xlsxs_names) {
      // let downloadName = "output-" + each[attrList['file']];
      const downloadName = "output-" + each["file"]
      //下载文件
      downloadSingleCode(downloadName)
      //添加到结果文件列表
      // addToResultList(each["file"])
    }
  }

  /**
   * 将文件添加到结果列表
   * @param {*} each
   */
  // function addToResultList(each: any) {
  //   //创建可以添加到结果列表的a标签
  //   const temp = document.createElement("a") as HTMLAnchorElement
  //   temp.id = each
  //   const resultListGroup = document.querySelector(".result-list-group") as any
  //   //在列表中添加该a标签
  //   resultListGroup.appendChild(temp)
  // }

  /**
   * 根据文件名称下载单个文件
   * @param {*} filenameParam
   */
  async function downloadSingleCode(filenameParam: any) {
    console.log("下载424编码.")
    const domain = "http://192.168.3.9:9092"
    // const domain = "http://127.0.0.1:9092"
    //调用后端下载文件接口
    console.log(domain + "/code/do424Download")
    const loadingInstance = ElLoading.service({ fullscreen: true })
    await axios
      .get(domain + "/code/do424Download", {
        headers: {
          "Access-Control-Allow-Origin": "*",
          "Cache-Control": "no-cache"
        },
        params: {
          filename: filenameParam
        },
        validateStatus: function (status) {
          return (status >= 200 && status < 300) || status == 304
        }
      })
      .then((res) => {
        loadingInstance.close()
        console.log(res)
        if (res.status == 200) {
          if (window.navigator && (window.navigator as any).msSaveOrOpenBlob) {
            // IE浏览器
            // window.navigator.msSaveOrOpenBlob(blob, fileName)
            // window.navigator.msSaveOrOpenBlob(blob)
          } else {
            //非IE浏览器 创建a标签模拟点击下载
            const downFile = document.createElement("a")
            downFile.style.display = "none"
            downFile.href = res.request.responseURL
            document.body.appendChild(downFile)
            downFile.click()
            document.body.removeChild(downFile) // 下载完成移除元素
          }
          ElMessage.success("下载成功")
        } else {
          ElMessage.error("下载失败，请稍后再试")
        }
        console.log("424下载程序结束.")
      })
      .catch((error) => {
        ElMessage.error("下载失败，请稍后再试")
        console.error(error)
        console.log("424下载程序结束.")
      })
  }

  /**
   * 在信息框中展示debug信息
   * @param {*} result_all_xlsxs
   */
  function showCodeDebugMsg(result_all_xlsxs: any) {
    let coding_msg_panel = document.querySelector(".coding-msg-panel") as any
    coding_msg_panel.innerHTML = ""
    for (const each of result_all_xlsxs) {
      for (const item of Object.values(each)) {
        const temp = document.createElement("div")
        temp.innerHTML += item
        temp.style.color = "black"
        coding_msg_panel = document.querySelector(".coding-msg-panel") as any
        //添加每个div的debug信息
        coding_msg_panel.appendChild(temp)
      }
    }
    //换行
    coding_msg_panel = document.querySelector(".coding-msg-panel") as any
    coding_msg_panel += "<br>"
  }
})
</script>

<template>
  <div class="four2four-container">
    <el-card>
      <!-- <router-view /> -->
      <div class="custom-page-title">操作命令框</div>

      <!-- operation panel -->

      <div class="operation-panel">
        <!-- button panel -->
        <div>
          <form @submit.prevent method="post" enctype="multipart/form-data" class="custom_form">
            <label class="el-button choose_file_button" for="source-files-uploader">请选择后缀为.xlsx的文件</label>
            <input
              type="file"
              name="sourceSheetFiles"
              class="el-button"
              id="source-files-uploader"
              accept=".xlsx"
              multiple
            />
            <input type="submit" class="el-button custom_button_submit" value="开始上传" />
          </form>
          <button type="button" class="el-button start_code_button">开始编码</button>
          <label class="el-button preview_code_button" for="result-files-uploader">请选择需要预览的Excel</label>
          <input
            type="file"
            name="resultSheetFiles"
            class="el-button"
            id="result-files-uploader"
            accept=".xlsx"
            multiple
          />
        </div>
      </div>

      <!-- coding msg panel -->

      <div class="code-info-panel">
        <div>
          <div class="code-msg-panel">
            <a disabled class="custom-page-title">编码运行信息</a>
            <div class="coding-msg-panel" />
          </div>
        </div>
      </div>

      <!--  source + excel -->

      <div class="information-panel">
        <!-- source list group -->
        <div class="file-list-panel">
          <div class="source-list-group">
            <a disabled class="source-list-group-start custom-page-title">源文件列表</a>
          </div>
        </div>

        <!-- Excel panel -->
        <div class="excel-panel">
          <div class="excel-panel-group">
            <a disabled class="custom-page-title">源文件Excel预览</a>
            <div id="source-excel-btns" />
            <div id="source-excel-panel" />
          </div>
        </div>
      </div>

      <!--  result + excel -->

      <div class="information-panel">
        <!-- result list group -->
        <div class="file-list-panel">
          <div class="result-list-group">
            <a disabled class="result-list-group-start custom-page-title">结果文件列表</a>
          </div>
        </div>

        <!-- Excel panel -->
        <div class="excel-panel">
          <div class="excel-panel-group">
            <a disabled class="custom-page-title">结果文件Excel预览</a>
            <div id="result-excel-btns" />
            <div id="result-excel-panel" />
          </div>
        </div>
      </div>
    </el-card>
  </div>
</template>

<style lang="scss" scoped>
.four2four-container {
  :deep(table) {
    border: 1px solid #0d0e10;
    border-collapse: collapse;
    overflow-y: scroll;
    overflow-x: scroll;
    height: 500px;
    display: block;
  }
  :deep(td) {
    border: 1px solid #0d0e10;
  }
  :deep(th) {
    border: 1px solid #0d0e10;
  }
  :deep(#result-files-uploader),
  :deep(#source-files-uploader) {
    display: none;
  }
  .custom_button_submit,
  .choose_file_button,
  .preview_code_button,
  .start_code_button {
    width: 188px;
    margin-left: 0px;
    padding-bottom: 5px;
    margin-bottom: 10px;
  }
  .operation-panel {
    margin-bottom: 20px;
    padding: 5px;
  }
  .code-info-panel {
    margin-bottom: 40px;
  }
  .information-panel {
    margin-top: 40px;
    margin-bottom: 20px;
  }
  .file-list-panel {
    margin-bottom: 30px;
  }
  .custom_form {
    display: inline-block;
  }
  .coding-msg-panel {
    overflow-y: scroll;
    overflow-x: scroll;
    height: 500px;
    background-color: rgb(238, 238, 238);
  }
  :deep(.custom-page-title) {
    font-size: 18px;
    font-weight: bold;
  }
}
</style>
