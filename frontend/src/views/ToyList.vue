<template>
  <div class="toy-list">
    <div class="operation-bar">
      <div class="search-form">
        <el-form :inline="true" :model="searchForm" ref="searchFormRef">
          <el-form-item label="品名" prop="name">
            <el-input v-model="searchForm.name" placeholder="请输入品名" clearable />
          </el-form-item>
          <el-form-item label="厂名" prop="factory_name">
            <el-input v-model="searchForm.factory_name" placeholder="请输入厂名" clearable />
          </el-form-item>
          <el-form-item label="货号" prop="factory_code">
            <el-input v-model="searchForm.factory_code" placeholder="请输入货号" clearable />
          </el-form-item>
          <el-form-item>
            <el-button type="primary" @click="handleSearch">搜索</el-button>
            <el-button @click="handleReset">重置</el-button>
          </el-form-item>
        </el-form>
      </div>
      <div class="button-group">
        <el-button type="primary" @click="showCreateDialog">新增货物</el-button>
        <el-button type="success" :disabled="!selectedItems.length" @click="handleExport">导出Excel</el-button>
        <el-button type="warning" @click="showImportDialog">导入Excel</el-button>
        <el-button type="danger" :disabled="!selectedItems.length" @click="batchDelete">批量删除</el-button>
      </div>
    </div>

    <el-table
      v-loading="loading"
      :data="items"
      @selection-change="handleSelectionChange"
      border
      style="width: 100%">
      <el-table-column type="selection" width="55" />
      <el-table-column label="图片" width="100">
        <template #default="{ row }">
          <el-image
            v-if="row.image_path"
            :src="`http://localhost:8000/${row.image_path}?t=${new Date().getTime()}`"
            fit="cover"
            style="width: 50px; height: 50px; cursor: pointer;"
            :preview-src-list="getPreviewImages(row)"
            :initial-index="0"
            preview-teleported
            @error="() => handleImageError(row)"
          />
          <el-icon v-else><Picture /></el-icon>
        </template>
      </el-table-column>
      <el-table-column prop="factory_code" label="货号" />
      <el-table-column prop="factory_name" label="厂名" />
      <el-table-column prop="name" label="品名" />
      <el-table-column prop="packaging" label="包装" />
      <el-table-column prop="packing_quantity" label="装箱量PCS" />
      <el-table-column prop="unit_price" label="单价" />
      <el-table-column prop="gross_weight" label="毛重KG" />
      <el-table-column prop="net_weight" label="净重KG" />
      <el-table-column prop="outer_box_size" label="外箱规格CM" />
      <el-table-column prop="product_size" label="产品规格" />
      <el-table-column prop="inner_box" label="内箱" />
      <el-table-column prop="remarks" label="备注" />
      <el-table-column prop="created_at" label="录入时间" width="150">
        <template #default="{ row }">
          {{ formatDateTime(row.created_at) }}
        </template>
      </el-table-column>
      <el-table-column prop="updated_at" label="更新时间" width="150">
        <template #default="{ row }">
          {{ formatDateTime(row.updated_at) }}
        </template>
      </el-table-column>
      <el-table-column label="操作" width="150">
        <template #default="{ row }">
          <el-button type="primary" link @click="showEditDialog(row)">编辑</el-button>
          <el-button type="danger" link @click="handleDelete(row)">删除</el-button>
        </template>
      </el-table-column>
    </el-table>

    <!-- 分页控件 -->
    <div class="pagination-container">
      <el-pagination
        v-model:current-page="pagination.currentPage"
        v-model:page-size="pagination.pageSize"
        :page-sizes="pagination.pageSizes"
        :total="pagination.total"
        layout="total, sizes, prev, pager, next, jumper"
        :prev-text="'上一页'"
        :next-text="'下一页'"
        :page-size-text="'条/页'"
        :total-text="'共 {total} 条'"
        :jumper-text="'前往'"
        :page-text="'页'"
        @size-change="handleSizeChange"
        @current-change="handleCurrentChange"
      />
    </div>

    <!-- 新增/编辑对话框 -->
    <el-dialog
      :title="dialogType === 'create' ? '新增货物' : '编辑货物'"
      v-model="dialogVisible"
      width="50%"
      @close="handleDialogClose">
      <el-form :model="form" :rules="rules" ref="formRef" label-width="120px">
        <el-form-item label="货号" prop="factory_code">
          <el-input v-model="form.factory_code" />
        </el-form-item>
        <el-form-item label="厂名" prop="factory_name">
          <el-input v-model="form.factory_name" />
        </el-form-item>
        <el-form-item label="品名" prop="name">
          <el-input v-model="form.name" />
        </el-form-item>
        <el-form-item label="包装" prop="packaging">
          <el-input v-model="form.packaging" />
        </el-form-item>
        <el-form-item label="装箱量PCS" prop="packing_quantity">
          <el-input-number v-model="form.packing_quantity" :min="1" />
        </el-form-item>
        <el-form-item label="单价" prop="unit_price">
          <el-input-number v-model="form.unit_price" :min="0" :precision="2" />
        </el-form-item>
        <el-form-item label="毛重KG" prop="gross_weight">
          <el-input-number v-model="form.gross_weight" :min="0" :precision="2" />
        </el-form-item>
        <el-form-item label="净重KG" prop="net_weight">
          <el-input-number v-model="form.net_weight" :min="0" :precision="2" />
        </el-form-item>
        <el-form-item label="外箱规格CM" prop="outer_box_size">
          <el-input v-model="form.outer_box_size" />
        </el-form-item>
        <el-form-item label="产品规格" prop="product_size">
          <el-input v-model="form.product_size" />
        </el-form-item>
        <el-form-item label="内箱" prop="inner_box">
          <el-input v-model="form.inner_box" />
        </el-form-item>
        <el-form-item label="备注" prop="remarks">
          <el-input v-model="form.remarks" type="textarea" :rows="2" />
        </el-form-item>
        <!-- 录入时间只在编辑模式下显示，且为只读 -->
        <el-form-item v-if="dialogType === 'edit'" label="录入时间">
          <el-input :value="formatDateTime(form.created_at)" readonly disabled />
        </el-form-item>
        <!-- 更新时间只在编辑模式下显示，且为只读 -->
        <el-form-item v-if="dialogType === 'edit'" label="更新时间">
          <el-input :value="formatDateTime(form.updated_at)" readonly disabled />
        </el-form-item>
        <el-form-item label="图片">
          <el-upload
            class="avatar-uploader"
            action="http://localhost:8000/items/"
            :show-file-list="false"
            :auto-upload="false"
            :on-change="handleImageChange"
            :before-upload="beforeImageUpload"
            :on-remove="handleImageRemove"
            :on-error="handleUploadError"
            accept="image/jpeg,image/png,image/gif,image/bmp"
            :limit="1"
            ref="imageUploadRef"
            :key="uploadKey">
            <img v-if="imageUrl" :src="imageUrl" class="avatar" />
            <el-icon v-else class="avatar-uploader-icon"><Plus /></el-icon>
            <template #tip>
              <div class="el-upload__tip">
                请上传图片文件(JPG、JPEG、PNG、GIF、BMP)
                <el-button v-if="imageUrl" type="primary" link @click="clearUploadImage">清除图片</el-button>
              </div>
            </template>
          </el-upload>
        </el-form-item>
      </el-form>
      <template #footer>
        <span class="dialog-footer">
          <el-button @click="dialogVisible = false">取消</el-button>
          <el-button type="primary" @click="handleSubmit">确定</el-button>
        </span>
      </template>
    </el-dialog>

    <!-- 导入对话框 -->
    <el-dialog
      title="导入Excel"
      v-model="importDialogVisible"
      width="400px">
      <el-form :model="importForm" ref="importFormRef" label-width="100px">
        <el-form-item label="厂名" prop="factory_name">
          <el-input v-model="importForm.factory_name" placeholder="请输入厂名" />
        </el-form-item>
        <el-form-item label="Excel文件" prop="file">
          <el-upload
            class="upload-demo"
            action="#"
            :auto-upload="false"
            :on-change="handleFileChange"
            :limit="1"
            ref="uploadRef"
            accept=".xlsx,.xls">
            <template #trigger>
              <el-button type="primary">选择文件</el-button>
            </template>
            <template #tip>
              <div class="el-upload__tip">
                请上传Excel文件(.xlsx, .xls)
                <el-button type="text" @click="downloadTemplate">下载模板</el-button>
              </div>
            </template>
          </el-upload>
        </el-form-item>

      </el-form>
      <template #footer>
        <span class="dialog-footer">
          <el-button @click="importDialogVisible = false">取消</el-button>
          <el-button type="primary" @click="handleImport" :loading="importing">导入</el-button>
        </span>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, onMounted, computed } from 'vue'
import { ElMessage, ElMessageBox, ElNotification } from 'element-plus'
import axios from 'axios'
import { Picture, Plus } from '@element-plus/icons-vue'

// 日期格式化函数
const formatDateTime = (dateString) => {
  if (!dateString) return ''
  const date = new Date(dateString)
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  const hours = String(date.getHours()).padStart(2, '0')
  const minutes = String(date.getMinutes()).padStart(2, '0')
  return `${year}-${month}-${day} ${hours}:${minutes}`
}

const loading = ref(false)
const items = ref([])
const selectedItems = ref([])
const dialogVisible = ref(false)
const dialogType = ref('create')
const imageUrl = ref('')
const imageFile = ref(null)
const imageUploadRef = ref(null)
const uploadKey = ref(Date.now()) // 用于强制重新创建上传组件

const form = ref({
  factory_code: '',
  factory_name: '',
  name: '',
  packaging: '',
  packing_quantity: 1,
  unit_price: 0,
  gross_weight: 0,
  net_weight: 0,
  outer_box_size: '',
  product_size: '',
  inner_box: '',
  remarks: ''
  // entry_time字段已移除，将在提交表单时自动设置
})

const searchForm = ref({
  name: '',
  factory_name: '',
  factory_code: ''
})

const searchFormRef = ref(null)

// 处理搜索
const handleSearch = async () => {
  // 搜索时重置为第一页
  pagination.value.currentPage = 1
  await fetchItems()
}

// 处理重置
const handleReset = () => {
  searchFormRef.value?.resetFields()
  fetchItems()
}

// 分页相关数据
const pagination = ref({
  currentPage: 1,
  pageSize: 10,
  total: 0,
  pageSizes: [10, 20, 50]
})

// 获取货物列表
const fetchItems = async () => {
  loading.value = true
  try {
    // 添加请求拦截器
    axios.interceptors.request.use(config => {
      console.log('请求发出:', config.method.toUpperCase(), config.url)
      return config
    })

    // 添加响应拦截器
    axios.interceptors.response.use(response => {
      console.log('响应收到:', {
        status: response.status,
        data: response.data
      })
      return response
    })

    const response = await axios.get('http://localhost:8000/items/', {
      params: {
        ...searchForm.value,
        page: pagination.value.currentPage,
        page_size: pagination.value.pageSize
      }
    })
    
    // 更新数据和分页信息
    items.value = response.data.items
    pagination.value.total = response.data.total
    pagination.value.currentPage = response.data.page
    pagination.value.pageSize = response.data.page_size
  } catch (error) {
    console.error('获取数据失败:', {
      message: error.message,
      response: error.response,
      config: error.config
    })
    ElMessage.error(`获取数据失败: ${error.response?.status || '网络连接异常'}`)
  } finally {
    loading.value = false
  }
}

// 处理表格选择
const handleSelectionChange = (selection) => {
  selectedItems.value = selection
}

// 处理页码变化
const handleCurrentChange = (page) => {
  pagination.value.currentPage = page
  fetchItems()
}

// 处理每页显示数量变化
const handleSizeChange = (size) => {
  pagination.value.pageSize = size
  // 切换每页数量时，如果当前页码超出范围，自动调整到第一页
  const maxPage = Math.ceil(pagination.value.total / size)
  if (pagination.value.currentPage > maxPage) {
    pagination.value.currentPage = 1
  }
  fetchItems()
}

// 显示创建对话框
const showCreateDialog = () => {
  dialogType.value = 'create'
  form.value = {
    factory_code: '',
    factory_name: '',
    name: '',
    packaging: '',
    packing_quantity: 1,
    unit_price: 0,
    gross_weight: 0,
    net_weight: 0,
    outer_box_size: '',
    product_size: '',
    inner_box: '',
    remarks: ''
    // entry_time字段已移除，将在提交表单时自动设置
  }
  imageUrl.value = ''
  imageFile.value = null
  // 重置上传组件
  if (imageUploadRef.value) {
    imageUploadRef.value.clearFiles()
  }
  // 更新uploadKey强制重新创建上传组件
  uploadKey.value = Date.now()
  console.log('打开创建对话框并更新uploadKey:', uploadKey.value)
  dialogVisible.value = true
}

// 显示编辑对话框
const showEditDialog = (row) => {
  dialogType.value = 'edit'
  // 创建一个新对象，确保数值字段被转换为数字类型
  form.value = { 
    ...row,
    // 显式转换数值字段为数字类型
    unit_price: Number(row.unit_price),
    packing_quantity: Number(row.packing_quantity),
    gross_weight: Number(row.gross_weight),
    net_weight: Number(row.net_weight)
  }
  // 添加时间戳参数防止浏览器缓存
  imageUrl.value = row.image_path ? `http://localhost:8000/${row.image_path}?t=${new Date().getTime()}` : ''
  imageFile.value = null
  // 重置上传组件
  if (imageUploadRef.value) {
    imageUploadRef.value.clearFiles()
  }
  // 更新uploadKey强制重新创建上传组件
  uploadKey.value = Date.now()
  console.log('打开编辑对话框并更新uploadKey:', uploadKey.value)
  dialogVisible.value = true
}

// 处理对话框关闭
const handleDialogClose = () => {
  // 释放之前创建的URL对象
  if (imageUrl.value && imageUrl.value.startsWith('blob:')) {
    URL.revokeObjectURL(imageUrl.value)
  }
  imageFile.value = null
  // 重置上传组件
  if (imageUploadRef.value) {
    imageUploadRef.value.clearFiles()
  }
  // 更新uploadKey强制重新创建上传组件
  uploadKey.value = Date.now()
  console.log('关闭对话框并更新uploadKey:', uploadKey.value)
}

// 验证上传的文件是否为图片类型
const beforeImageUpload = (file) => {
  const isImage = file.type.startsWith('image/');
  const isAllowedType = ['image/jpeg', 'image/png', 'image/gif', 'image/bmp'].includes(file.type);
  
  if (!isImage || !isAllowedType) {
    ElMessage.error('只能上传图片文件(JPG、JPEG、PNG、GIF、BMP)！');
    return false;
  }
  return true;
}

// 处理图片变更
const handleImageChange = (file) => {
  console.log('handleImageChange 触发', file)
  console.log('当前文件状态:', {
    fileExists: !!file,
    fileRawExists: file && !!file.raw,
    currentImageUrl: imageUrl.value,
    currentImageFile: imageFile.value ? '存在' : '不存在'
  })
  
  if (file && file.raw) {
    // 再次验证文件类型
    if (beforeImageUpload(file.raw)) {
      console.log('文件类型验证通过，准备更新图片')
      // 如果之前有创建的URL对象，先释放它以避免内存泄漏
      if (imageUrl.value && imageUrl.value.startsWith('blob:')) {
        console.log('释放之前的blob URL:', imageUrl.value)
        URL.revokeObjectURL(imageUrl.value)
      }
      
      // 强制重置上传组件状态
      if (imageUploadRef.value) {
        console.log('重置上传组件状态')
        imageUploadRef.value.clearFiles()
      }
      
      imageFile.value = file.raw
      imageUrl.value = URL.createObjectURL(file.raw)
      console.log('已更新图片:', {
        newImageUrl: imageUrl.value,
        fileType: file.raw.type,
        fileSize: file.raw.size
      })
      
      // 更新uploadKey强制重新创建上传组件
      uploadKey.value = Date.now()
      console.log('更新uploadKey:', uploadKey.value)
      
      // 强制更新上传组件的显示
      setTimeout(() => {
        if (imageUploadRef.value) {
          console.log('触发上传组件重新渲染')
          imageUploadRef.value.$forceUpdate()
        }
      }, 0)
    } else {
      console.log('文件类型验证失败')
    }
  } else {
    console.log('没有收到有效的文件对象')
  }
}

// 清除已上传的图片
const clearUploadImage = () => {
  // 释放之前创建的URL对象
  if (imageUrl.value && imageUrl.value.startsWith('blob:')) {
    URL.revokeObjectURL(imageUrl.value)
  }
  imageUrl.value = ''
  imageFile.value = null
  // 重置上传组件
  if (imageUploadRef.value) {
    imageUploadRef.value.clearFiles()
  }
  // 更新uploadKey强制重新创建上传组件
  uploadKey.value = Date.now()
  console.log('清除图片并更新uploadKey:', uploadKey.value)
}

// 处理图片移除
const handleImageRemove = (file) => {
  console.log('handleImageRemove 触发', file)
  // 释放之前创建的URL对象
  if (imageUrl.value && imageUrl.value.startsWith('blob:')) {
    console.log('释放blob URL:', imageUrl.value)
    URL.revokeObjectURL(imageUrl.value)
  }
  imageUrl.value = ''
  imageFile.value = null
  console.log('图片已移除')
}

// 处理上传错误
const handleUploadError = (err, file, fileList) => {
  console.log('上传错误:', err)
  console.log('错误文件:', file)
  console.log('文件列表:', fileList)
  ElMessage.error(`上传失败: ${err.message || '未知错误'}`)
}

// 获取预览图片列表
const getPreviewImages = (row) => {
  if (!row.image_path) return []
  // 添加时间戳参数防止浏览器缓存
  return [`http://localhost:8000/${row.image_path}?t=${new Date().getTime()}`]
}

// 处理图片加载错误
const handleImageError = (row) => {
  console.error('图片加载失败:', row.image_path)
  ElNotification({
    title: '图片加载失败',
    message: `无法加载图片: ${row.name || row.factory_code}`,
    type: 'warning',
    duration: 3000
  })
}

const formRef = ref(null)
const rules = {
  factory_code: [
    { required: true, message: '请输入货号', trigger: 'blur' },
    { min: 2, max: 100, message: '货号长度应在2-100个字符之间', trigger: 'blur' }
  ],
  factory_name: [
    { required: true, message: '请输入厂名', trigger: 'blur' },
    { min: 2, max: 64, message: '厂名长度应在2-64个字符之间', trigger: 'blur' }
  ],
  name: [
    { required: true, message: '请输入品名', trigger: 'blur' },
    { min: 2, max: 100, message: '品名长度应在2-100个字符之间', trigger: 'blur' }
  ],
  packaging: [
    { required: true, message: '请输入包装', trigger: 'blur' },
    { min: 2, max: 100, message: '包装长度应在2-100个字符之间', trigger: 'blur' }
  ],
  packing_quantity: [
    { required: true, message: '请输入装箱量PCS', trigger: 'blur' },
    { type: 'number', min: 1, message: '装箱量PCS必须大于0', trigger: 'blur' }
  ],
  unit_price: [
    { required: true, message: '请输入单价', trigger: 'blur' },
    { type: 'number', min: 0, message: '单价不能为负数', trigger: 'blur' }
  ],
  gross_weight: [
    { required: true, message: '请输入毛重KG', trigger: 'blur' },
    { type: 'number', min: 0, message: '毛重不能为负数', trigger: 'blur' }
  ],
  net_weight: [
    { required: true, message: '请输入净重KG', trigger: 'blur' },
    { type: 'number', min: 0, message: '净重不能为负数', trigger: 'blur' },
    { validator: (rule, value, callback) => {
      if (value > form.value.gross_weight) {
        callback(new Error('净重不能大于毛重'))
      } else {
        callback()
      }
    }, trigger: 'blur' }
  ],
  outer_box_size: [
    { required: true, message: '请输入外箱规格CM', trigger: 'blur' }
  ],
  product_size: [
    { required: true, message: '请输入产品规格', trigger: 'blur' }
  ],
  inner_box: [
    { required: true, message: '请输入内箱', trigger: 'blur' }
  ]
}

// 处理表单提交
const handleSubmit = async () => {
  if (!formRef.value) return
  
  try {
    await formRef.value.validate()
    const formData = new FormData()
    
    // 确保数值类型字段被正确转换
    formData.append('factory_code', form.value.factory_code)
    formData.append('factory_name', form.value.factory_name)
    formData.append('name', form.value.name)
    formData.append('packaging', form.value.packaging)
    formData.append('packing_quantity', form.value.packing_quantity.toString())
    formData.append('unit_price', form.value.unit_price.toString())
    formData.append('gross_weight', form.value.gross_weight.toString())
    formData.append('net_weight', form.value.net_weight.toString())
    formData.append('outer_box_size', form.value.outer_box_size)
    formData.append('product_size', form.value.product_size)
    formData.append('inner_box', form.value.inner_box)
    formData.append('remarks', form.value.remarks || '')
    
    // 创建新货物时自动使用当前时间作为录入时间
    if (dialogType.value === 'create') {
      formData.append('entry_time', new Date().toISOString())
    } else if (form.value.entry_time) {
      // 编辑时保留原有录入时间
      formData.append('entry_time', form.value.entry_time)
    }
    
    // 编辑模式下保留原有的created_at（录入时间）
    if (dialogType.value === 'edit' && form.value.created_at) {
      formData.append('created_at', form.value.created_at)
    }
    
    if (imageFile.value) {
      formData.append('image', imageFile.value)
    }

    if (dialogType.value === 'create') {
      await axios.post('http://localhost:8000/items/', formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        }
      })
      ElMessage.success('创建成功')
    } else {
      await axios.put(`http://localhost:8000/items/${form.value.id}`, formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        }
      })
      ElMessage.success('更新成功')
    }
    dialogVisible.value = false
    fetchItems()
  } catch (error) {
    console.error('操作失败:', error)
    
    // 更详细的错误处理
    if (error.response?.data?.detail) {
      // 后端返回的详细错误信息
      ElMessage.error(`操作失败: ${error.response.data.detail}`)
    } else if (error.response?.status === 422 && error.response?.data?.detail) {
      // 处理验证错误
      const validationErrors = error.response.data.detail
      const errorMessages = validationErrors.map(err => `${err.loc[1]}: ${err.msg}`).join('\n')
      ElMessage.error(`数据验证失败:\n${errorMessages}`)
    } else if (error.message) {
      // 一般错误信息
      ElMessage.error(`操作失败: ${error.message}`)
    } else if (typeof error === 'string') {
      // 字符串形式的错误
      ElMessage.error(`操作失败: ${error}`)
    } else {
      // 未知错误
      ElMessage.error('操作失败：请检查输入是否正确')
    }
  }
}

// 处理删除
const handleDelete = async (row) => {
  try {
    await ElMessageBox.confirm('确定要删除这条记录吗？')
    await axios.delete(`http://localhost:8000/items/${row.id}`)
    ElMessage.success('删除成功')
    fetchItems()
  } catch (error) {
    if (error !== 'cancel') {
      ElMessage.error('删除失败')
    }
  }
}

// 批量删除
const batchDelete = async () => {
  if (selectedItems.value.length === 0) {
    ElMessage.warning('请选择要删除的货物')
    return
  }
  
  try {
    await ElMessageBox.confirm(`确定要删除选中的 ${selectedItems.value.length} 条记录吗？`)
    await axios.delete('http://localhost:8000/items/batch', {
      data: {
        item_ids: selectedItems.value.map(item => item.id)
      }
    })
    ElMessage.success('批量删除成功')
    fetchItems()
    // 清空选中状态
    selectedItems.value = []
  } catch (error) {
    if (error !== 'cancel') {
      console.error('批量删除失败:', error)
      ElMessage.error('批量删除失败')
    }
  }
}

// 处理导出
const handleExport = async () => {
  if (selectedItems.value.length === 0) {
    ElMessage.warning('请选择要导出的货物')
    return
  }
  try {
    const response = await axios.post('http://localhost:8000/items/export', {
      item_ids: selectedItems.value.map(item => item.id)
    }, {
      responseType: 'blob'
    })
    
    // 创建下载链接
    const url = window.URL.createObjectURL(new Blob([response.data]))
    const link = document.createElement('a')
    link.href = url
    link.setAttribute('download', `货物报价表_${new Date().getTime()}.xlsx`)
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
    
    ElMessage.success('导出成功')
  } catch (error) {
    ElMessage.error('导出失败')
  }
}

// 导入相关
const importDialogVisible = ref(false)
const importForm = ref({
  factory_name: '',
  file: null
})
const importFormRef = ref(null)
const uploadRef = ref(null)
const importing = ref(false)

// 显示导入对话框
const showImportDialog = () => {
  importDialogVisible.value = true
  importForm.value = {
    factory_name: '',
    file: null
  }
}

// 处理文件选择
const handleFileChange = (file) => {
  importForm.value.file = file.raw
}

// 处理导入
const handleImport = async () => {
  if (!importForm.value.file) {
    ElMessage.warning('请选择Excel文件')
    return
  }
  
  importing.value = true
  try {
    const formData = new FormData()
    formData.append('file', importForm.value.file)
    formData.append('factory_name', importForm.value.factory_name)
    
    // 使用唯一的导入方法
    const importEndpoint = 'http://localhost:8000/items/import'
    
    console.log('使用导入方法')
    const response = await axios.post(importEndpoint, formData)
    
    if (response.data && response.data.imported_count) {
      ElMessage.success(`成功导入 ${response.data.imported_count} 条数据`)
      importDialogVisible.value = false
      fetchItems()
      importForm.value = {
        factory_name: '',
        file: null
      }
      if (importFormRef.value) {
        importFormRef.value.resetFields()
      }
      if (uploadRef.value) {
        uploadRef.value.clearFiles()
      }
    } else {
      ElMessage.error('导入失败：未导入任何数据')
    }
  } catch (error) {
    console.error('导入失败:', error)
    
    // 获取详细的错误信息
    let errorMessage = '导入失败'
    if (error.response?.data?.detail) {
      errorMessage = error.response.data.detail
    } else if (error.message) {
      errorMessage = error.message
    }
    
    // 显示详细的错误信息
    ElMessage({
      message: errorMessage,
      type: 'error',
      duration: 8000, // 延长显示时间以便用户阅读完整错误信息
      showClose: true
    })
    
    // 如果是数据转换错误，给出额外提示
    if (errorMessage.includes('could not convert') || errorMessage.includes('导入第') || errorMessage.includes('行时出错')) {
      setTimeout(() => {
        ElMessage({
          message: '提示：请检查Excel文件中的数值字段格式，确保数值字段只包含数字，不包含文字说明',
          type: 'warning',
          duration: 10000,
          showClose: true
        })
      }, 1000)
    }
  } finally {
    importing.value = false
  }
}

// 下载导入模板
const downloadTemplate = async () => {
  try {
    const response = await axios.get('http://localhost:8000/items/import-template', {
      responseType: 'blob'
    })
    
    // 创建下载链接
    const url = window.URL.createObjectURL(new Blob([response.data]))
    const link = document.createElement('a')
    link.href = url
    link.setAttribute('download', '货物导入模板.xlsx')
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    window.URL.revokeObjectURL(url)
    
    ElMessage.success('模板下载成功')
  } catch (error) {
    ElMessage.error('模板下载失败')
  }
}

onMounted(() => {
  fetchItems()
})
</script>

<style scoped>
.toy-list {
  padding: 20px;
}

.operation-bar {
  display: flex;
  justify-content: space-between;
  margin-bottom: 20px;
}

.button-group {
  display: flex;
  gap: 10px;
}

.pagination-container {
  display: flex;
  justify-content: center;
  margin-top: 20px;
  margin-bottom: 20px;
}

.avatar-uploader .avatar {
  width: 178px;
  height: 178px;
  display: block;
}

.avatar-uploader .el-upload {
  border: 1px dashed var(--el-border-color);
  border-radius: 6px;
  cursor: pointer;
  position: relative;
  overflow: hidden;
  transition: var(--el-transition-duration-fast);
}

.avatar-uploader .el-upload:hover {
  border-color: var(--el-color-primary);
}

.avatar-uploader-icon {
  font-size: 28px;
  color: #8c939d;
  width: 178px;
  height: 178px;
  text-align: center;
}

.image-preview {
  display: flex;
  justify-content: center;
  align-items: center;
  height: 100%;
}

.image-preview img {
  max-width: 100%;
  max-height: 100%;
}

.preview-dialog .el-dialog__body {
  height: 70vh;
  display: flex;
  justify-content: center;
  align-items: center;
}

.search-form {
  margin-bottom: 20px;
}

.table-operations {
  margin-bottom: 20px;
}

.table-operations .el-button {
  margin-right: 10px;
}

.el-table .warning-row {
  --el-table-tr-bg-color: var(--el-color-warning-light-9);
}

.el-table .success-row {
  --el-table-tr-bg-color: var(--el-color-success-light-9);
}

.el-table .error-row {
  --el-table-tr-bg-color: var(--el-color-error-light-9);
}

.table-header-operations {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 20px;
}

.table-header-operations .left-operations {
  display: flex;
  gap: 10px;
}

.table-header-operations .right-operations {
  display: flex;
  gap: 10px;
}

.export-options {
  margin-top: 10px;
}

.export-options .el-checkbox {
  margin-right: 15px;
  margin-bottom: 10px;
}

.export-options-title {
  margin-bottom: 10px;
  font-weight: bold;
}
</style>