from abc import ABC, abstractmethod
from sqlalchemy.orm import Session
from src.models import Check
from src.repository.base import BaseRepository
from src.repository.models import CheckInfoModel, CheckItemModel
# from sqlalchemy import func
from src.models import Item
from datetime import datetime
from decimal import Decimal


class BaseDatabaseRepository(BaseRepository, ABC):
    """Базовый класс для репозиториев, использующих SQLAlchemy."""
    
    @abstractmethod
    def _get_session(self) -> Session:
        """Получает сессию SQLAlchemy."""
        ...
    
    def append_check(self, check_id: int, check: Check):
        """Добавляет чек в базу данных."""
        
        
        with self._get_session() as session:
            existing = session.query(CheckInfoModel).get(check_id)
            if existing:
                self._update_check_model(existing, check)
            else:
                check_model = self._create_check_model(check_id, check)
                session.add(check_model)
            
            session.commit()
    
    def _create_check_model(self, check_id: int, check: Check) -> 'CheckInfoModel':
        """Создает модель чека из бизнес-объекта."""

        check_model = CheckInfoModel(
            id=check_id,
            msg_type=check.msg_type,
            address1=check.address1,
            address2=check.address2,
            date=check.date.isoformat(timespec="minutes"),
            cashier=check.cashier,
            total=check.total
        )
        
        check_model.items = [
            CheckItemModel(
                msg_type=check.msg_type,
                product_name=item.product_name,
                price=item.price,
                qty=item.qty,
                amount=item.amount,
                uom=item.uom
            )
            for item in check.items
        ]
        
        return check_model
    
    def _update_check_model(self, model: 'CheckInfoModel', check: Check):
        """Обновляет существующую модель чека."""
        
        
        model.msg_type = check.msg_type
        model.address1 = check.address1
        model.address2 = check.address2
        model.date = check.date.isoformat(timespec="minutes")
        model.cashier = check.cashier
        model.total = check.total
        
        # Удаляем старые позиции и добавляем новые
        model.items.clear()
        model.items = [
            CheckItemModel(
                msg_type=check.msg_type,
                product_name=item.product_name,
                price=item.price,
                qty=item.qty,
                amount=item.amount,
                uom=item.uom
            )
            for item in check.items
        ]
    
    def get_check(self, check_id: int) -> Check:
        """Получает чек по ID."""

        
        with self._get_session() as session:
            check_model = session.query(CheckInfoModel).get(check_id)
            if not check_model:
                return None
            
            items = [
                Item(
                    product_name=item.product_name,
                    price=Decimal(str(item.price)),
                    qty=Decimal(str(item.qty)),
                    amount=Decimal(str(item.amount)),
                    uom=item.uom
                )
                for item in check_model.items
            ]
            
            return Check(
                msg_type=check_model.msg_type,
                address1=check_model.address1,
                address2=check_model.address2,
                date=datetime.fromisoformat(check_model.date),
                cashier=check_model.cashier,
                total=Decimal(str(check_model.total)),
                items=items
            )
    
    # def get_stats(self) -> dict:
    #     """Получает статистику из БД."""

    #     with self._get_session() as session:
    #         stats = {
    #             'total_checks': session.query(func.count(CheckInfoModel.id)).scalar(),
    #             'total_items': session.query(func.count(CheckItemModel.id)).scalar(),
    #             'total_amount': session.query(func.sum(CheckInfoModel.total)).scalar() or 0,
    #         }
            
    #         type_stats = (
    #             session.query(
    #                 CheckInfoModel.msg_type,
    #                 func.count(CheckInfoModel.id)
    #             )
    #             .group_by(CheckInfoModel.msg_type)
    #             .all()
    #         )
    #         stats['checks_by_type'] = dict(type_stats)
            
    #         return stats