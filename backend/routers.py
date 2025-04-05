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
from import_service import import_items

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
async def import_excel(file: UploadFile = File(...), factory_name: str = Form(...), db: Session = Depends(models.get_db)):
    # 调用import_service中的import_items函数
    return await import_items(file=file, factory_name=factory_name, db=db)

# 导入Excel数据路由已统一，使用唯一的导入方法

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