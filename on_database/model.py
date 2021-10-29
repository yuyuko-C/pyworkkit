import peewee as pw
import pandas as pd
from on_database.select import ModelSelect

class BaseModel(pw.Model):

    @classmethod
    def login_mysql(cls, db_name: str, host: str, port: int, user: str, password: str):
        db = pw.MySQLDatabase(db_name,
                              host=host,
                              port=port,
                              user=user,
                              password=password)
        cls._meta.database = db

    @classmethod
    def login_sqlite(cls, db_name: str, host: str, port: int, user: str, password: str):
        db = pw.SqliteDatabase(db_name,
                               host=host,
                               port=port,
                               user=user,
                               password=password)
        cls._meta.database = db

    @classmethod
    def login_postgre(cls, db_name: str, host: str, port: int, user: str, password: str):
        db = pw.PostgresqlDatabase(db_name,
                                   host=host,
                                   port=port,
                                   user=user,
                                   password=password)
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
    def exists(cls):
        db: pw.Database = cls._meta.database
        return db.table_exists(cls.__name__.lower())

    @classmethod
    def create(cls):
        db: pw.Database = cls._meta.database
        with db.atomic():
            db.create_tables([cls])

    @classmethod
    def drop(cls):
        db: pw.Database = cls._meta.database
        with db.atomic():
            db.drop_tables([cls])

    @classmethod
    def reinstance(cls):
        if cls.exists():
            cls.drop()
        cls.create()

    @classmethod
    def append_df(cls, df: pd.DataFrame, fields: list = None):
        df = df.fillna('')
        fields = fields if fields else df.columns.tolist()
        cls.insert_many(df.values, fields).execute()

    @classmethod
    def rewrite_df(cls, df: pd.DataFrame):
        cls.reinstance()
        cls.append_df(df)

    @classmethod
    def to_dataframe(cls):
        return pd.DataFrame(cls.select().dicts())

    @classmethod
    def reset_auto_increment(cls):
        db: pw.Database = cls._meta.database
        table_name = cls.__name__.lower()
        db.execute_sql("alter table `{}` auto_increment=1;".format(table_name))

    @classmethod
    def select(cls, *fields):
        is_default = not fields
        if not fields:
            fields = cls._meta.sorted_fields
        return ModelSelect(cls, fields, is_default=is_default)