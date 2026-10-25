import os
from sqlalchemy import create_engine
from sqlalchemy.orm import scoped_session, sessionmaker

import config

engine = create_engine(config.database_url(), pool_pre_ping=True, future=True)
Session = scoped_session(sessionmaker(bind=engine, autoflush=True, expire_on_commit=False))


def migrate():

    from alembic import command
    from alembic.config import Config
    cfg = Config(os.path.join(config.BASE, "alembic.ini"))
    cfg.set_main_option("script_location", os.path.join(config.BASE, "alembic"))
    command.upgrade(cfg, "head")
