from sqlalchemy import Column, Integer, String, Float, DateTime, ForeignKey, create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime
import os

# 创建数据库连接
DATABASE_URL = "sqlite:///./toy_management.db"  # 开发阶段使用SQLite，生产环境可替换为MySQL
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()

# 定义货物报价表模型
class ToyItem(Base):
    __tablename__ = "toy_items"
    
    id = Column(Integer, primary_key=True, index=True)
    factory_code = Column(String(100), comment="货号")
    name = Column(String(100), comment="品名")
    packaging = Column(String(100), comment="包装")
    packing_quantity = Column(Integer, comment="装箱量PCS")
    unit_price = Column(Float, comment="单价")
    gross_weight = Column(Float, comment="毛重KG")
    net_weight = Column(Float, comment="净重KG")
    outer_box_size = Column(String(100), comment="外箱规格CM")
    product_size = Column(String(100), comment="产品规格")
    inner_box = Column(String(100), comment="内箱")
    remarks = Column(String(255), comment="备注")
    image_path = Column(String(255), comment="图片路径")
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)
    
    def __repr__(self):
        return f"<ToyItem {self.factory_code}: {self.name}>"

# 创建数据库表
def create_tables():
    # 如果数据库文件存在，先删除它
    if os.path.exists("toy_management.db"):
        try:
            os.remove("toy_management.db")
            print("已删除旧的数据库文件，将创建新的数据库。")
        except Exception as e:
            print(f"删除数据库文件失败: {e}")
    
    # 创建所有表
    Base.metadata.create_all(bind=engine)
    print("数据库表创建成功！")

# 获取数据库会话
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()