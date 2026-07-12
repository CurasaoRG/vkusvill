from sqlalchemy import create_engine, Column, Integer, String, Numeric, DateTime, ForeignKey, Index
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from datetime import datetime

Base = declarative_base()


class CheckInfoModel(Base):
    __tablename__ = 'check_info'
    
    id = Column(Integer, primary_key=True)
    msg_type = Column(String, nullable=False)
    address1 = Column(String)
    address2 = Column(String)
    date = Column(String, nullable=False)
    cashier = Column(String)
    total = Column(Numeric(10, 2), nullable=False)
    created_at = Column(DateTime, default=datetime.datetime.now())
    
    items = relationship("CheckItemModel", back_populates="check", cascade="all, delete-orphan")
    
    __table_args__ = (
        Index('idx_check_info_date', 'date'),
        Index('idx_check_info_msg_type', 'msg_type'),
    )


class CheckItemModel(Base):
    __tablename__ = 'check_items'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    check_id = Column(Integer, ForeignKey('check_info.id', ondelete='CASCADE'), nullable=False)
    msg_type = Column(String, nullable=False)
    product_name = Column(String, nullable=False)
    price = Column(Numeric(10, 2), nullable=False)
    qty = Column(Numeric(10, 3), nullable=False)
    amount = Column(Numeric(10, 2), nullable=False)
    uom = Column(String)
    created_at = Column(DateTime, default=datetime.datetime.now())
    
    check = relationship("CheckInfoModel", back_populates="items")
    
    __table_args__ = (
        Index('idx_check_items_check_id', 'check_id'),
    )