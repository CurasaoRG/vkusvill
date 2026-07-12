from pathlib import Path
from contextlib import contextmanager
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, Session
from src.repository.database_repo import BaseDatabaseRepository
from src.repository.models import Base


class SQLiteRepository(BaseDatabaseRepository):
    """Репозиторий для работы с SQLite через SQLAlchemy."""
    
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.engine = create_engine(
            f'sqlite:///{db_path}',
            echo=False, 
            connect_args={'check_same_thread': False}  
        )
        self.SessionLocal = sessionmaker(bind=self.engine)
        Base.metadata.create_all(self.engine)
    
    @contextmanager
    def _get_session(self) -> Session:
        """Предоставляет сессию SQLAlchemy."""
        session = self.SessionLocal()
        try:
            yield session
        except Exception:
            session.rollback()
            raise
        finally:
            session.close()