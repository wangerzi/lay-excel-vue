<template>
  <div>
    <p>
      简单使用下导出，<b>点我导出</b> 按钮会以 xlsx 格式导出 Hello World，<b>导出样式</b> 按钮会以 xlsx 格式导出带样式的表格。
    </p>
    <p>
      更多 <b>数据导出，边框、行高列宽、样式配置、多 Sheet 支持</b>请见文档
      <a target="_blank" href="http://excel.wj2015.com/_book/">http://excel.wj2015.com/_book/</a>
    </p>
    <p>
      样式相关函数文档：<a target="_blank" href="http://excel.wj2015.com/_book/docs/%E5%87%BD%E6%95%B0%E5%88%97%E8%A1%A8/%E6%A0%B7%E5%BC%8F%E8%AE%BE%E7%BD%AE%E7%9B%B8%E5%85%B3%E5%87%BD%E6%95%B0.html">样式设置相关函数</a>
    </p>
    <p>有任何疑问或 BUG 反馈，请加QQ群 <b>555056599</b>，或邮件 <a href="mailto:admin@wj2015.com">admin@wj2015.com</a>，Github 可能响应不及时</p>
    <p>
      <button @click="handleExport">点我导出</button>
      <button @click="handleExportStyle">导出样式</button>
    </p>
    <div>
      <form action="" @submit="handleSubmit">
        <p>
          <input type="file" name="excel" id="upload-excel-file">
          <a href="http://excel.wj2015.com/demos/test_import.xlsx" target="_blank">下载样例文件</a>
        </p>
        <p>
          <button type="submit">解析excel</button>
        </p>
      </form>
    </div>
  </div>
</template>

<script>
import LAY_EXCEL from 'lay-excel';
export default {
  name: 'HelloWorld',
  methods: {
    handleSubmit(e) {
      e.preventDefault();

      let files = document.getElementById('upload-excel-file').files;
      console.log(files);
      try {
        LAY_EXCEL.importExcel(files, {}, function (data) {
          alert('导入JSON：' + JSON.stringify(data));
          data = LAY_EXCEL.filterImportData(data, {
            'id': 'A'
            , 'username': 'B'
            , 'experience': 'C'
            , 'sex': 'D'
            , 'score': 'E'
            , 'city': 'F'
            , 'classify': 'G'
            , 'wealth': 'H'
            , 'sign': 'I'
          })
          alert('梳理后JSON：' + JSON.stringify(data));
          console.log(data);
        });
      } catch (e) {
        alert('导入错误: ' + e.message);
      }
    },
    handleExport() {
      LAY_EXCEL.exportExcel([['Hello', 'World', '!']], 'hello.xlsx', 'xlsx');
    },
    handleExportStyle() {
      let data = [
        ['A1', 'B2', 'C3', 'D4'],
        ['A1', 'B2', 'C3', 'D4'],
        ['A1', 'B2', 'C3', 'D4'],
      ];
      LAY_EXCEL.setExportCellStyle(data, 'B2:C3', {
        s: {
          fill: { bgColor: { indexed: 64 }, fgColor: { rgb: "FF0000" } },
          alignment: {
            horizontal: 'center',
            vertical: 'center'
          }
        }
      });
      LAY_EXCEL.exportExcel(data, 'hello_style.xlsx', 'xlsx');
    }
  }
}
</script>

