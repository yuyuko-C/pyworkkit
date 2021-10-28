import peewee as pw
import pandas as pd

class ModelSelect(pw.ModelSelect):
    
    @classmethod
    def to_dataframe(cls):
        return pd.DataFrame(cls.dicts())
