from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form, Body, Response
from starlette.responses import FileResponse
from sqlalchemy.orm import Session
from typing import List
import models
from datetime import datetime
import os
import shutil
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from io import BytesIO
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask

router = APIRouter()

# 配置图片上传目录
UPLOAD_DIR = "uploads"
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
        file_path = os.path.join(UPLOAD_DIR, file_name)
        
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(image.file, buffer)
        image_path = file_path
    
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
        if db_item.image_path and os.path.exists(db_item.image_path):
            os.remove(db_item.image_path)
        
        # 保存新图片
        file_ext = os.path.splitext(image.filename)[1]
        file_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}{file_ext}"
        file_path = os.path.join(UPLOAD_DIR, file_name)
        
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(image.file, buffer)
        image_path = file_path
    
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
    if db_item.image_path and os.path.exists(db_item.image_path):
        os.remove(db_item.image_path)
    
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
            # 使用PIL打开并调整图片大小，保持EXIF方向信息
            pil_image = PILImage.open(item.image_path)
            # 获取EXIF信息并保持图片方向
            try:
                for orientation in pil_image._getexif().values():
                    if orientation == 3:
                        pil_image = pil_image.rotate(180, expand=True)
                    elif orientation == 6:
                        pil_image = pil_image.rotate(270, expand=True)
                    elif orientation == 8:
                        pil_image = pil_image.rotate(90, expand=True)
            except (AttributeError, KeyError, IndexError, TypeError):
                # 如果没有EXIF信息，保持原样
                pass
            
            # 保持原始图片大小
            
            # 将PIL图片转换为openpyxl可用的格式
            img_byte_arr = BytesIO()
            pil_image.save(img_byte_arr, format=pil_image.format if pil_image.format else 'PNG')
            img_byte_arr.seek(0)
            
            # 创建openpyxl图片对象并插入到单元格
            img = Image(img_byte_arr)
            # 调整单元格大小以适应原始图片
            ws.row_dimensions[row].height = 150  # 设置更大的行高以适应原始图片
            ws.column_dimensions['A'].width = 30  # 设置图片列宽度
            
            # 将图片添加到工作表并定位到对应单元格
            img.anchor = f'A{row}'
            ws.add_image(img)
    
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
                    # 图片字段在Excel中只是一个提示，实际导入时不会处理图片数据
                    image_path=None
                )
                
                # 如果提供了厂名但Excel中没有厂名字段，使用表单提供的厂名
                if not new_item.factory_name:
                    new_item.factory_name = factory_name
                
                # 检查必填字段
                if not new_item.factory_code or not new_item.name or not new_item.packaging:
                    continue  # 跳过不完整的行
                
                # 添加到数据库
                db.add(new_item)
                imported_count += 1
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