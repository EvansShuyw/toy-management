import { createApp } from 'vue'
import { createRouter, createWebHistory } from 'vue-router'
import { createPinia } from 'pinia'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import App from './App.vue'
import ToyList from './views/ToyList.vue'

// 创建路由配置
const routes = [
  { path: '/', redirect: '/toys' },
  { path: '/toys', component: ToyList }
]

const router = createRouter({
  history: createWebHistory(),
  routes
})

// 创建Pinia状态管理
const pinia = createPinia()

// 创建Vue应用实例
const app = createApp(App)

// 使用插件
app.use(router)
app.use(pinia)
app.use(ElementPlus)

// 挂载应用
app.mount('#app')