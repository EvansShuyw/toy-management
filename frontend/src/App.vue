<template>
  <el-container class="app-container">
    <el-header>
      <h1>货物报价管理系统</h1>
    </el-header>
    <el-main>
      <router-view></router-view>
    </el-main>
  </el-container>
</template>

<script setup>
import { ElContainer, ElHeader, ElMain } from 'element-plus'
</script>

<style>
/* 全局样式 */
html, body {
  font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
}

/* 设置Element Plus组件的全局样式 */
:root {
  --el-font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
  --el-font-size-base: 16px;
  --el-font-weight-primary: bold;
}

.el-table {
  font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
  font-size: 16px;
  font-weight: bold;
}

.el-table th, .el-table td {
  text-align: center;
}

.el-form-item__label, .el-input__inner, .el-button, .el-select, .el-dropdown, .el-menu-item {
  font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
  font-size: 16px;
  font-weight: bold;
}

.el-dialog__title, .el-message-box__title {
  font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
  font-size: 18px;
  font-weight: bold;
  text-align: center;
}
</style>

<style scoped>
.app-container {
  height: 100vh;
}

.el-header {
  background-color: #409EFF;
  color: white;
  display: flex;
  align-items: center;
  justify-content: center;
}

.el-header h1 {
  font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
  font-size: 24px;
  font-weight: bold;
}

.el-main {
  padding: 20px;
  background-color: #f5f7fa;
}
</style>