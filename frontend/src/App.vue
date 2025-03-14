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

<style scoped>
.app-container {
  height: 100vh;
}

.el-header {
  background-color: #409EFF;
  color: white;
  display: flex;
  align-items: center;
}

.el-main {
  padding: 20px;
  background-color: #f5f7fa;
}
</style>