<template>
  <div>
    <p>
      简单使用下导出，点击如下按钮会以 xlsx 格式导出 Hello World
    </p>
    <p>
      更多使用介绍请见文档
      <a target="_blank" href="http://excel.wj2015.com/_book/">http://excel.wj2015.com/_book/</a>
    </p>
    <button @click="handleExport">点我导出</button>
  </div>
</template>

<script>
import LAY_EXCEL from 'lay-excel';
export default {
  name: 'HelloWorld',
  methods: {
    handleExport() {
      LAY_EXCEL.downloadExl([['Hello', 'World', '!']], 'hello', 'xlsx');
    }
  }
}
</script>

