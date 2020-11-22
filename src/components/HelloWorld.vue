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
    <button @click="handleExport">点我导出</button>
    <button @click="handleExportStyle">导出样式</button>
  </div>
</template>

<script>
import LAY_EXCEL from 'lay-excel';
export default {
  name: 'HelloWorld',
  methods: {
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

