from fastapi import UploadFile, File, Form, HTTPException, Depends
from sqlalchemy.orm import Session
import models
from datetime import datetime
import os
import shutil
import uuid
from openpyxl import load_workbook
from PIL import Image as PILImage
from io import BytesIO
import zipfile
import re

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")

async def import_items(
    file: UploadFile = File(...),
    factory_name: str = Form(...),
    db: Session = Depends(models.get_db)
):
    # 检查文件类型
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="只支持Excel文件格式(.xlsx, .xls)")

    # 保存上传的文件到临时位置
    temp_file_path = f"uploads/temp_{file.filename}"
    with open(temp_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        # 使用新方法处理Excel文件
        # 1. 将.xlsx文件当作zip文件打开，直接提取其中的图片
        images = []
        if file.filename.endswith('.xlsx'):
            try:
                # 尝试从xlsx文件中提取图片
                with zipfile.ZipFile(temp_file_path, 'r') as zip_ref:
                    # 查找所有图片文件
                    for item in zip_ref.namelist():
                        if item.startswith('xl/media/') and item.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                            # 提取图片数据
                            image_data = zip_ref.read(item)
                            images.append(image_data)
                            print(f"从Excel中提取图片: {item}")
            except Exception as e:
                print(f"提取Excel图片时出错: {str(e)}")
                # 继续处理，即使没有图片

        # 2. 使用openpyxl读取数据
        workbook = load_workbook(temp_file_path)
        sheet = workbook.active

        # 获取表头（第一行）
        headers = [cell.value for cell in sheet[1]]

        # 定义字段映射（Excel列名 -> 数据库字段名）
        field_mapping = {
            "图片": "image_path",
            "货号": "factory_code",
            "厂名": "factory_name",
            "品名": "name",
            "包装": "packaging",
            "装箱量PCS": "packing_quantity",
            "单价": "unit_price",
            "毛重KG": "gross_weight",
            "净重KG": "net_weight",
            "外箱规格CM": "outer_box_size",
            "产品规格": "product_size",
            "内箱": "inner_box",
            "备注": "remarks"
        }

        # 找出每个字段在Excel中的列索引
        field_indices = {}
        for i, header in enumerate(headers):
            if header in field_mapping:
                field_indices[field_mapping[header]] = i

        # 检查必填字段是否存在
        required_fields = ["factory_code", "name", "packaging", "packing_quantity"]
        missing_fields = [field for field in required_fields if field not in field_indices]
        if missing_fields:
            missing_headers = [key for key, value in field_mapping.items() if value in missing_fields]
            raise HTTPException(status_code=400, detail=f"Excel文件缺少必要的列: {', '.join(missing_headers)}")

        # 导入数据
        imported_count = 0
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2), start=2):  # 从第二行开始（跳过表头）
            try:
                # 处理图片
                image_path = None
                if "image_path" in field_indices and images:
                    # 计算当前行对应的图片索引
                    img_index = row_idx - 2  # 从0开始
                    
                    # 如果有对应的图片，处理并保存
                    if img_index < len(images):
                        try:
                            # 获取图片数据
                            image_data = images[img_index]
                            
                            # 验证文件是否为图片格式
                            try:
                                # 从图片数据创建PIL Image对象
                                img = PILImage.open(BytesIO(image_data))
                                
                                # 如果是RGBA模式，转换为RGB
                                if img.mode in ('RGBA', 'LA'):
                                    background = PILImage.new('RGB', img.size, (255, 255, 255))
                                    background.paste(img, mask=img.split()[-1])
                                    img = background
                            except Exception as e:
                                print(f"无效的图片格式: {str(e)}")
                                continue  # 跳过无效的图片
                            
                            # 生成唯一的文件名
                            file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}.png"
                            os.makedirs(UPLOAD_DIR, exist_ok=True)
                            file_path = os.path.abspath(os.path.join(UPLOAD_DIR, file_name))
                            
                            # 确保文件路径在uploads目录内
                            if not os.path.commonpath([file_path, os.path.abspath(UPLOAD_DIR)]) == os.path.abspath(UPLOAD_DIR):
                                raise ValueError("无效的文件路径")
                            
                            # 保存为PNG格式
                            img.save(file_path, 'PNG', optimize=True, quality=100)
                            print(f'成功保存图片到：{file_path}')
                            
                            # 设置为相对URL路径以便前端访问
                            image_path = f"uploads/{file_name}"
                            print(f'图片路径已设置: {image_path}')
                        except Exception as e:
                            print(f"处理第 {row_idx} 行图片时出错: {str(e)}")

                # 创建新记录
                new_item = models.ToyItem(
                    factory_code=row[field_indices.get("factory_code")].value if "factory_code" in field_indices else None,
                    factory_name=row[field_indices.get("factory_name")].value if "factory_name" in field_indices else factory_name,
                    name=row[field_indices.get("name")].value if "name" in field_indices else None,
                    packaging=row[field_indices.get("packaging")].value if "packaging" in field_indices else None,
                    packing_quantity=int(row[field_indices.get("packing_quantity")].value) if "packing_quantity" in field_indices and row[field_indices.get("packing_quantity")].value is not None else 0,
                    unit_price=float(row[field_indices.get("unit_price")].value) if "unit_price" in field_indices and row[field_indices.get("unit_price")].value is not None else 0.0,
                    gross_weight=float(row[field_indices.get("gross_weight")].value) if "gross_weight" in field_indices and row[field_indices.get("gross_weight")].value is not None else 0.0,
                    net_weight=float(row[field_indices.get("net_weight")].value) if "net_weight" in field_indices and row[field_indices.get("net_weight")].value is not None else 0.0,
                    outer_box_size=row[field_indices.get("outer_box_size")].value if "outer_box_size" in field_indices else None,
                    product_size=row[field_indices.get("product_size")].value if "product_size" in field_indices else None,
                    inner_box=row[field_indices.get("inner_box")].value if "inner_box" in field_indices else None,
                    remarks=row[field_indices.get("remarks")].value if "remarks" in field_indices else None,
                    image_path=image_path
                )

                # 如果提供了厂名但Excel中没有厂名字段，使用表单提供的厂名
                if not new_item.factory_name:
                    new_item.factory_name = factory_name

                # 检查必填字段
                missing_fields = []
                if not new_item.factory_code:
                    missing_fields.append("货号")
                if not new_item.name:
                    missing_fields.append("品名")
                if not new_item.packaging:
                    missing_fields.append("包装")

                if missing_fields:
                    print(f"导入第 {row_idx} 行时缺少必填字段: {', '.join(missing_fields)}")
                    continue

                try:
                    # 添加到数据库
                    db.add(new_item)
                    imported_count += 1
                except Exception as e:
                    print(f"导入第 {row_idx} 行时数据库操作失败: {str(e)}")
                    continue
            except Exception as e:
                # 记录错误但继续处理其他行
                print(f"导入第 {row_idx} 行时出错: {str(e)}")
                continue

        # 提交事务
        db.commit()

        # 返回导入结果
        return {"imported_count": imported_count, "message": "导入成功"}

    except Exception as e:
        # 回滚事务
        db.rollback()
        raise HTTPException(status_code=500, detail=f"导入失败: {str(e)}")

    finally:
        # 清理临时文件
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)