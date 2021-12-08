import peewee as pw


class Database:
    @classmethod
    def login_mysql(cls, db_name: str, host: str, port: int, user: str, password: str):
        db = pw.MySQLDatabase(db_name,
                              host=host,
                              port=port,
                              user=user,
                              password=password)
        return db

    @classmethod
    def login_sqlite(cls, db_name: str, host: str, port: int, user: str, password: str):
        db = pw.SqliteDatabase(db_name,
                               host=host,
                               port=port,
                               user=user,
                               password=password)
        return db

    @classmethod
    def login_postgre(cls, db_name: str, host: str, port: int, user: str, password: str):
        db = pw.PostgresqlDatabase(db_name,
                                   host=host,
                                   port=port,
                                   user=user,
                                   password=password)
        return db

