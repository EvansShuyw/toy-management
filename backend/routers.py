from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form, Body, Response
from starlette.responses import FileResponse
from sqlalchemy.orm import Session
from typing import List
import models
from datetime import datetime
import os
import shutil
import uuid
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from io import BytesIO
import openpyxl.styles
import os.path
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask

router = APIRouter()

# 配置图片上传目录
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 获取所有货物报价表项目
@router.get("/items/", response_model=List[dict])
def get_items(name: str = None, factory_name: str = None, db: Session = Depends(models.get_db)):
    query = db.query(models.ToyItem)
    if name:
        query = query.filter(models.ToyItem.name.ilike(f"%{name}%"))
    if factory_name:
        query = query.filter(models.ToyItem.factory_name.ilike(f"%{factory_name}%"))
    items = query.all()
    return [{
        "id": item.id,
        "factory_code": item.factory_code,
        "factory_name": item.factory_name,
        "name": item.name,
        "packaging": item.packaging,
        "packing_quantity": item.packing_quantity,
        "unit_price": item.unit_price,
        "gross_weight": item.gross_weight,
        "net_weight": item.net_weight,
        "outer_box_size": item.outer_box_size,
        "product_size": item.product_size,
        "inner_box": item.inner_box,
        "remarks": item.remarks,
        "image_path": item.image_path,
        "created_at": item.created_at,
        "updated_at": item.updated_at
    } for item in items]

# 创建新的货物报价表项目
@router.post("/items/")
async def create_item(factory_code: str = Form(...),
                     factory_name: str = Form(...),
                     name: str = Form(...),
                     packaging: str = Form(...),
                     packing_quantity: int = Form(...),
                     unit_price: float = Form(...),
                     gross_weight: float = Form(...),
                     net_weight: float = Form(...),
                     outer_box_size: str = Form(...),
                     product_size: str = Form(...),
                     inner_box: str = Form(...),
                     remarks: str = Form(None),
                     image: UploadFile = File(None),
                     db: Session = Depends(models.get_db)):
    # 处理图片上传
    image_path = None
    if image:
        file_ext = os.path.splitext(image.filename)[1]
        file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}{file_ext}"
        os.makedirs(UPLOAD_DIR, exist_ok=True)
    try:
        file_path = os.path.join(UPLOAD_DIR, file_name)
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(image.file, buffer)
        # 确保文件路径正确
        image_path = os.path.relpath(file_path, os.path.dirname(__file__)).replace('\\', '/')
        # 设置为相对URL路径以便前端访问
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"文件保存失败: {str(e)}")
    
    # 创建数据库记录
    db_item = models.ToyItem(
        factory_code=factory_code,
        factory_name=factory_name,
        name=name,
        packaging=packaging,
        packing_quantity=packing_quantity,
        unit_price=unit_price,
        gross_weight=gross_weight,
        net_weight=net_weight,
        outer_box_size=outer_box_size,
        product_size=product_size,
        inner_box=inner_box,
        remarks=remarks,
        image_path=image_path
    )
    db.add(db_item)
    db.commit()
    db.refresh(db_item)
    return db_item

# 更新货物报价表项目
@router.put("/items/{item_id}")
async def update_item(item_id: int,
                     factory_code: str = Form(...),
                     factory_name: str = Form(...),
                     name: str = Form(...),
                     packaging: str = Form(...),
                     packing_quantity: int = Form(...),
                     unit_price: float = Form(...),
                     gross_weight: float = Form(...),
                     net_weight: float = Form(...),
                     outer_box_size: str = Form(...),
                     product_size: str = Form(...),
                     inner_box: str = Form(...),
                     remarks: str = Form(None),
                     image: UploadFile = File(None),
                     db: Session = Depends(models.get_db)):
    # 查找要更新的记录
    db_item = db.query(models.ToyItem).filter(models.ToyItem.id == item_id).first()
    if not db_item:
        raise HTTPException(status_code=404, detail="Item not found")
    
    # 处理图片上传
    image_path = db_item.image_path
    if image:
        # 删除旧图片
        if db_item.image_path:
            old_file_path = os.path.join(os.path.dirname(__file__), db_item.image_path.lstrip('/'))
            if os.path.exists(old_file_path):
                os.remove(old_file_path)
        
        # 保存新图片
        file_ext = os.path.splitext(image.filename)[1]
        file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}{file_ext}"
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        try:
            file_path = os.path.join(UPLOAD_DIR, file_name)
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(image.file, buffer)
            # 设置为相对URL路径以便前端访问
            image_path = f"uploads/{file_name}"
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"文件保存失败: {str(e)}")
    
    # 更新记录
    db_item.factory_code = factory_code
    db_item.factory_name = factory_name
    db_item.name = name
    db_item.packaging = packaging
    db_item.packing_quantity = packing_quantity
    db_item.unit_price = unit_price
    db_item.gross_weight = gross_weight
    db_item.net_weight = net_weight
    db_item.outer_box_size = outer_box_size
    db_item.product_size = product_size
    db_item.inner_box = inner_box
    db_item.remarks = remarks
    db_item.image_path = image_path
    db_item.updated_at = datetime.now()
    
    db.commit()
    db.refresh(db_item)
    return db_item

# 删除货物报价表项目
@router.delete("/items/{item_id}")
def delete_item(item_id: int, db: Session = Depends(models.get_db)):
    db_item = db.query(models.ToyItem).filter(models.ToyItem.id == item_id).first()
    if not db_item:
        raise HTTPException(status_code=404, detail="Item not found")
    
    # 删除关联的图片文件
    if db_item.image_path:
        file_path = os.path.join(UPLOAD_DIR, os.path.basename(db_item.image_path))
        if os.path.exists(file_path):
            os.remove(file_path)
    
    db.delete(db_item)
    db.commit()
    return {"message": "Item deleted successfully"}

# 导出选中的货物报价表为Excel
@router.post("/items/export")
async def export_items(request: dict = Body(...), db: Session = Depends(models.get_db)):
    item_ids = request.get("item_ids", [])
    # 获取选中的项目
    items = db.query(models.ToyItem).filter(models.ToyItem.id.in_(item_ids)).all()
    if not items:
        raise HTTPException(status_code=400, detail="No items found for export")
    
    # 创建Excel工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "货物报价表"
    
    # 添加表头
    headers = ["图片", "货号", "厂名", "品名", "包装", "装箱量PCS", "单价", "毛重KG", "净重KG", "外箱规格CM", "产品规格", "内箱", "备注"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    # 添加数据
    for row, item in enumerate(items, 2):
        ws.cell(row=row, column=2, value=item.factory_code)
        ws.cell(row=row, column=3, value=item.factory_name)
        ws.cell(row=row, column=4, value=item.name)
        ws.cell(row=row, column=5, value=item.packaging)
        ws.cell(row=row, column=6, value=item.packing_quantity)
        ws.cell(row=row, column=7, value=item.unit_price)
        ws.cell(row=row, column=8, value=item.gross_weight)
        ws.cell(row=row, column=9, value=item.net_weight)
        ws.cell(row=row, column=10, value=item.outer_box_size)
        ws.cell(row=row, column=11, value=item.product_size)
        ws.cell(row=row, column=12, value=item.inner_box)
        ws.cell(row=row, column=13, value=item.remarks)
        
        # 处理图片
        if item.image_path and os.path.exists(item.image_path):
            # 使用PIL打开并转换图片
            pil_image = PILImage.open(item.image_path)
            # 转换图片为RGB模式（如果是RGBA）
            if pil_image.mode in ('RGBA', 'LA'):
                background = PILImage.new('RGB', pil_image.size, (255, 255, 255))
                background.paste(pil_image, mask=pil_image.split()[-1])
                pil_image = background
            
            # 获取原始图片尺寸
            img_width, img_height = pil_image.size
            
            # 设置更合理的尺寸限制，保持更高的图片质量
            max_height = 600  # 增加最大高度限制
            max_width = 800   # 增加最大宽度限制
            
            # 计算缩放比例，但保持更高的图片质量
            scale = min(max_width/img_width, max_height/img_height, 1.0)
            new_width = int(img_width * scale)
            new_height = int(img_height * scale)
            
            # 仅当图片超过最大限制时才调整大小，使用高质量的缩放算法
            if scale < 1.0:
                pil_image = pil_image.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
            
            # 保存为临时的BytesIO对象，使用最高质量设置
            img_byte_arr = BytesIO()
            pil_image.save(img_byte_arr, format='PNG', optimize=False, quality=100)
            img_byte_arr.seek(0)
            
            # 创建openpyxl图片对象
            img = Image(img_byte_arr)
            
            # 根据图片实际大小设置单元格，增加系数以确保单元格足够大
            row_height = new_height * 0.85  # 增加Excel单元格高度转换因子
            col_width = new_width * 0.18   # 增加Excel单元格宽度转换因子
            
            # 设置单元格大小
            ws.row_dimensions[row].height = row_height
            ws.column_dimensions['A'].width = col_width
            
            # 将图片添加到单元格，使用精确定位
            cell_address = f'A{row}'
            ws.add_image(img, cell_address)
            # 调整单元格对齐方式
            cell = ws.cell(row=row, column=1)
            cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

    
    # 保存Excel文件
    file_name = f"货物报价表_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    os.makedirs("exports", exist_ok=True)
    file_path = os.path.join("exports", file_name)
    wb.save(file_path)
    
    # 返回文件并在发送后删除
    return FileResponse(
        path=file_path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=BackgroundTask(lambda: os.unlink(file_path))
    )

# 导入Excel数据
@router.post("/items/import")
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
        # 打开Excel文件
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
                if "image_path" in field_indices:
                    try:
                        # 获取单元格中的图片
                        cell_coord = f"{chr(65 + field_indices['image_path'])}{row_idx}"
                        # 使用_images属性获取工作表中的所有图片
                        for image in sheet._images:
                            if hasattr(image, 'anchor') and hasattr(image.anchor, '_from') and image.anchor._from.coord == cell_coord:
                                try:
                                    # 获取图片数据
                                    image_data = image._data()
                                    if not image_data:
                                        print(f"第 {row_idx} 行的图片数据为空")
                                        continue
                                    
                                    # 从图片数据创建PIL Image对象
                                    img = PILImage.open(BytesIO(image_data))
                                    
                                    # 如果是RGBA模式，转换为RGB
                                    if img.mode in ('RGBA', 'LA'):
                                        background = PILImage.new('RGB', img.size, (255, 255, 255))
                                        background.paste(img, mask=img.split()[-1])
                                        img = background
                                    
                                    # 生成唯一的文件名
                                    file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}.png"
                                    os.makedirs(UPLOAD_DIR, exist_ok=True)
                                    file_path = os.path.abspath(os.path.join(UPLOAD_DIR, file_name))
                                    
                                    # 确保文件路径在uploads目录内
                                    if not os.path.commonpath([file_path, os.path.abspath(UPLOAD_DIR)]) == os.path.abspath(UPLOAD_DIR):
                                        raise ValueError("无效的文件路径")
                                    
                                    # 保存为PNG格式，使用最高质量设置
                                    img.save(file_path, 'PNG', optimize=True, quality=100)
                                    print(f'成功保存图片到：{file_path}')
                                    # 设置为相对URL路径以便前端访问
                                    image_path = f"uploads/{file_name}"
                                    print(f'图片路径已设置: {image_path}')
                                    break  # 找到并处理了图片后就退出循环
                                except Exception as e:
                                    print(f"处理第 {row_idx} 行图片时出错: {str(e)}")
                                    continue
                    except Exception as e:
                        print(f"处理Excel图片时出错: {str(e)}")
                        # 继续处理其他数据，不中断导入过程

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

# 获取Excel导入模板
@router.get("/items/import-template")
async def get_import_template():
    # 创建Excel工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "货物导入模板"
    
    # 添加表头
    headers = ["图片", "货号", "厂名", "品名", "包装", "装箱量PCS", "单价", "毛重KG", "净重KG", "外箱规格CM", "产品规格", "内箱", "备注"]
    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)
    
    # 设置列宽
    for col_idx in range(1, len(headers) + 1):
        ws.column_dimensions[chr(64 + col_idx)].width = 15
    
    # 特别设置图片列的宽度
    ws.column_dimensions['A'].width = 20
    
    # 保存到临时文件
    template_path = "exports/import_template.xlsx"
    os.makedirs("exports", exist_ok=True)
    wb.save(template_path)
    
    # 返回文件
    return FileResponse(
        path=template_path,
        filename="货物导入模板.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=BackgroundTask(lambda: os.unlink(template_path))
    )