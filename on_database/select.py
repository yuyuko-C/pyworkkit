import peewee as pw
import pandas as pd

class ModelSelect(pw.ModelSelect):

    def to_dataframe(self):
        return pd.DataFrame(self.dicts())

    def where(self, *expressions):
        ret: ModelSelect = super().where(*expressions)
        return ret

    def with_cte(self, *cte_list):
        ret: ModelSelect = super().with_cte(*cte_list)
        return ret

    def where(self, *expressions):
        ret: ModelSelect = super().where(*expressions)
        return ret

    def orwhere(self, *expressions):
        ret: ModelSelect = super().orwhere(*expressions)
        return ret

    def order_by(self, *values):
        ret: ModelSelect = super().order_by(*values)
        return ret

    def order_by_extend(self, *values):
        ret: ModelSelect = super().order_by_extend(*values)
        return ret

    def limit(self, value=None):
        ret: ModelSelect =  super().limit(value=value)
        return ret

    def offset(self, value=None):
        ret: ModelSelect =  super().offset(value=value)
        return ret

    def paginate(self, page, paginate_by=20):
        ret: ModelSelect =  super().paginate(page, paginate_by=paginate_by)
        return ret

    def join(self, dest, join_type=..., on=None, src=None, attr=None):
        ret: ModelSelect = super().join(dest, join_type, on, src, attr)
        return ret