import typing
import peewee as pw
import pandas as pd
from .select import ModelSelect




class BaseModel(pw.Model):

    @classmethod
    def get_database(cls)->typing.Union[pw.Database,None]:
        return cls._meta.database 

    @classmethod
    def set_database(cls,db:pw.Database):
        cls._meta.database = db

    @classmethod
    def reconnect(cls, method):
        def inner(*args, **kargs):
            db: pw.Database = cls._meta.database
            if db.is_closed():
                db.connect()
            method(*args, **kargs)
            return
        return inner
    
    @classmethod
    def recreate(cls):
        if cls.table_exists():
            cls.drop_table()
        cls.create_table()

    @classmethod
    def append_df(cls, df: pd.DataFrame, fields: list = None):
        df = df.fillna('')
        fields = fields if fields else df.columns.tolist()
        cls.insert_many(df.values, fields).execute()

    @classmethod
    def rewrite_df(cls, df: pd.DataFrame):
        cls.recreate()
        cls.append_df(df)

    @classmethod
    def to_dataframe(cls):
        return pd.DataFrame(cls.select().dicts())

    @classmethod
    def reset_auto_increment(cls):
        db = cls.get_database()
        table_name = cls.__name__.lower()
        db.execute_sql("alter table `{}` auto_increment=1;".format(table_name))

    @classmethod
    def select(cls, *fields):
        is_default = not fields
        if not fields:
            fields = cls._meta.sorted_fields
        return ModelSelect(cls, fields, is_default=is_default)
