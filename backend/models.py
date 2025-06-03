from sqlalchemy import Column, Integer, String, Float, DateTime, ForeignKey, create_engine, Numeric
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime
import os

# 创建数据库连接
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_URL = f"sqlite:///{os.path.join(BASE_DIR, 'toy_management.db')}"
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

# 定义货物报价表模型
class ToyItem(Base):
    __tablename__ = "toy_items"
    
    id = Column(Integer, primary_key=True, index=True)
    factory_code = Column(String(100), comment="货号")
    factory_name = Column(String(64), comment="厂名")
    name = Column(String(100), comment="品名")
    packaging = Column(String(100), comment="包装")
    packing_quantity = Column(Integer, comment="装箱量PCS")
    unit_price = Column(Numeric(10, 3), comment="单价")
    gross_weight = Column(Float, comment="毛重KG")
    net_weight = Column(Float, comment="净重KG")
    outer_box_size = Column(String(100), comment="外箱规格CM")
    product_size = Column(String(100), comment="产品规格")
    inner_box = Column(String(100), comment="内箱")
    remarks = Column(String(255), comment="备注")
    image_path = Column(String(255), comment="图片路径")
    origin_sheet = Column(String(255), default='', comment='来源工作表名称')
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    
    def __repr__(self):
        return f"<ToyItem {self.factory_code}: {self.name}>"

# 创建数据库表
def create_tables():
    # 检查数据库文件是否存在
    db_path = os.path.join(BASE_DIR, 'toy_management.db')
    db_exists = os.path.exists(db_path)
    
    # 创建所有表（如果不存在）
    Base.metadata.create_all(bind=engine)
    
    if not db_exists:
        print("数据库文件创建成功！")
    else:
        print("数据库连接正常！")

# 获取数据库会话
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()