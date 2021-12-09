import pandas as pd
from datetime import date, timedelta
import calendar
from itertools import product
from typing import *

__all__=[
    "Time",
    "GroupBy",
    "AssistCol",
    "Others",
]


class Time:

    @classmethod
    def next_month_date(cls, cur_date: date, step: int):
        '''
        得到一年内的下一个月头的日期
        '''
        if (step > 12) or (step < -12):
            raise ValueError('months is out of range')

        cur_year, cur_month = cur_date.year, cur_date.month

        month_sum = cur_month+step
        if month_sum > 12:
            next_year = cur_year+1
            next_month = month_sum-12
        elif month_sum < 1:
            next_year = cur_year-1
            next_month = month_sum+12
        else:
            next_year = cur_year
            next_month = month_sum

        return date(next_year, next_month, 1)

    @classmethod
    def month_days(cls, year: int, month: int):
        '''
        计算该日期所在的月份总共有多少天
        '''
        weekday, monthdays = calendar.monthrange(year, month)
        return monthdays

    @classmethod
    def month_progress(cls, cut_date: date):
        '''
        计算该日期所在的月份的月时间进度
        '''
        return cut_date.day/cls.month_days(cut_date.year,cut_date.month)

    @classmethod
    def get_time_progress(cls, date: date, cut_date: date):
        if cut_date.year > date.year:
            return 1
        elif cut_date.year == date.year:
            if cut_date.month > date.month:
                return 1
            elif cut_date.month == date.month:
                return cls.month_progress(cut_date)
        else:
            return 0

    @classmethod
    def time_progress(cls, df: pd.DataFrame, idx_list: list, drop_na: bool):

        def get_dataframe_part(df: pd.DataFrame):

            df['time_progress'] = df['deliver_date'].apply(
                cls.get_time_progress, args=(cls.get_cut_date(df),))
            pt = df.pivot_table(
                'time_progress', idx_list, 'month_assist')
            columns = []
            columns.extend(idx_list)
            columns.extend(pt.columns)

            row_list = pt.values.tolist()
            lst = [pd.Series(row).dropna().tolist() for row in row_list]
            length = [len(row) for row in lst]
            value = row_list[length.index(max(length))]

            return value, columns

        if drop_na == False:
            time_lst, columns = get_dataframe_part(df)
            ret = pd.DataFrame(columns=columns)
            idx_lists = [df[idx].drop_duplicates().tolist()
                         for idx in idx_list]
            for i in product(*idx_lists):
                row = list(i)
                row.extend(time_lst)
                df = pd.DataFrame(row, index=columns)
                ret = ret.append(df.T)

            ret = ret.set_index(idx_list, drop=True).astype(float)
        else:
            df['time_progress'] = df['deliver_date'].apply(
                cls.get_time_progress, args=(cls.get_cut_date(df),))
            ret = df.pivot_table(
                'time_progress', idx_list, 'month_assist')

        return ret
    
    @staticmethod
    def get_cut_date(df: pd.DataFrame) -> date:
        date_col = [col for col in df.columns if 'date' in col][0]
        m_date = df[date_col].max()
        return m_date


class GroupBy:

    @staticmethod
    def fill_dic(keys: list, value_list: list):
        if len(keys) != len(value_list):
            raise ValueError('length is not equal')

        dic = {}
        length = len(keys)

        for i in range(length):
            dic[keys[i]] = value_list[i]

        return dic

    @staticmethod
    def group_by(df: pd.DataFrame, by: Union[str, list], aggfunc: str) -> pd.DataFrame:
        '''[summary]

        Args:
            df (pd.DataFrame): The DataFrame to group by.
            by (Union[str, list]): Can use 'all' or the str or list in DataFrame columns or index name.
            aggfunc (str): Can use 'sum','mean','min','max','count's

        Returns:
            pd.DataFrame: The DataFrame after group by.
        '''
        if aggfunc == 'sum':
            if by == 'all':
                df = pd.DataFrame(df.sum()).T
            else:
                df = df.groupby(by).sum()
        elif aggfunc == 'mean':
            if by == 'all':
                df = pd.DataFrame(df.mean()).T
            else:
                df = df.groupby(by).mean()
        elif aggfunc == 'min':
            if by == 'all':
                df = pd.DataFrame(df.min()).T
            else:
                df = df.groupby(by).min()
        elif aggfunc == 'max':
            if by == 'all':
                df = pd.DataFrame(df.max()).T
            else:
                df = df.groupby(by).max()
        elif aggfunc == 'count':
            if by == 'all':
                df = pd.DataFrame(df.count()).T
            else:
                df = df.groupby(by).count()
        return df

    @staticmethod
    def dicts_group_by(dic: dict, by: Union[str, list], aggfunc: str) -> Dict[str,pd.DataFrame]:
        '''[summary]

        Args:
            dic (dict): DataFrame dict
            by (Union[str, list]): The str or list in DataFrame columns or index name.
            aggfunc (str): Can use 'sum','mean','min','max','count'

        Returns:
            dict: A dict saved the DataFrame after group_by in the same key.
        '''
        m_dic = {}
        keys = list(dic.keys())
        for key in keys:
            df = dic[key].copy(True)
            if isinstance(df, pd.DataFrame):
                if by == None:
                    m_dic[key] = df
                else:
                    df = GroupBy.group_by(df, by, aggfunc)
                    m_dic[key] = df

        return m_dic

    @staticmethod
    def fixed_index(df: pd.DataFrame, idx_list: list) -> pd.DataFrame:
        '''Set DataFrame index as idx_list.Not exisit index will be filled with 'All'.

        Args:
            df (pd.DataFrame): [description].
            idx_list (list): [description]

        Returns:
            pd.DataFrame: [description]
        '''

        # 补全index
        df.reset_index(inplace=True)

        for col in idx_list:
            if col not in df.columns:
                df[col] = 'All'

        df = df.set_index(idx_list)

        if 'index' in df.columns:
            df.drop('index', axis=1, inplace=True)
        return df

    @staticmethod
    def multi_group_by(df: pd.DataFrame, idx_list: list, aggfunc: str) -> List[pd.DataFrame]:
        '''Group_by df by 'all',idx_list,source and concat them to an empty DataFrame.

        Args:
            df (pd.DataFrame): [description]
            idx_list (list): [description]
            aggfunc (str): [description]

        Returns:
            pd.DataFrame: [description]
        '''
        df = df.copy(True)
        ret = []

        # 所有的合计计算
        all_all = GroupBy.group_by(df, 'all', aggfunc)
        all_all = GroupBy.fixed_index(all_all, idx_list)
        ret.append(all_all)

        # 部分的合计计算
        if len(idx_list) > 1:
            part = [GroupBy.group_by(
                df, i, aggfunc) for i in idx_list]
            part = [GroupBy.fixed_index(p, idx_list) for p in part]
            ret.extend(part)

        # 传入的本体
        ret.append(df)

        return ret


class AssistCol:
    @staticmethod
    def day_assist_col(df: pd.DataFrame):

        date_col = [col for col in df.columns if 'date' in col][0]
        day_series = df[date_col].apply(lambda x: x.strftime('%YY%m%d'))

        return day_series

    @staticmethod
    def week_assist_col(df: pd.DataFrame):

        date_col = [col for col in df.columns if 'date' in col][0]

        save_str = ['']

        def date_to_week(x: date):
            weekday = x.weekday()
            if (save_str[0] == '') | (weekday == 0):
                monday_date = x-timedelta(weekday)
                sunday_date = monday_date+timedelta(days=6)
                monday_date = monday_date.strftime('%YY%m%d')
                sunday_date = sunday_date.strftime('%YY%m%d')
                year_week = '{}-{}'.format(monday_date, sunday_date)
                save_str[0] = year_week
            return save_str[0]

        week_series = df[date_col].apply(date_to_week)

        return week_series

    @staticmethod
    def month_assist_col(df: pd.DataFrame):

        # 月辅助列
        date_col = [col for col in df.columns if 'date' in col][0]
        month_series = df[date_col].apply(lambda x: x.strftime('%Y年%m月'))

        return month_series

    @staticmethod
    def area_assist_col(df: pd.DataFrame):
        # 地区辅助列
        area_series = df['customer_id'].apply(
            lambda x: '深圳' if 'cn' in x else '杭州')
        return area_series


class Others:
    @classmethod
    def str_fix(cls, ustring: str):
        if not isinstance(ustring, str):
            if ustring != None:
                return ustring
            else:
                return 'None'

        def strQ2B(ustring):
            """全角转半角"""
            rstring = ""
            for uchar in ustring:
                inside_code = ord(uchar)
                if inside_code == 12288:  # 全角空格直接转换
                    inside_code = 32
                elif (inside_code >= 65281 and inside_code <= 65374):  # 全角字符（除空格）根据关系转化
                    inside_code -= 65248

                rstring += chr(inside_code)
            return rstring

        ustring = strQ2B(ustring)
        return ustring.strip().lower().replace('\n', '')

    @classmethod
    def data_fix(cls, data):
        if isinstance(data, (int, float)):
            return data
        elif isinstance(data, str):
            data = data.strip()
            return float(data)
        else:
            raise TypeError('data type is {}'.format(type(data)))

    @classmethod
    def pivot2df(cls, df: pd.DataFrame, index_name: list, columns: list):
        ret = pd.DataFrame()
        if df.index.names != [None]:
            if (df.index.names != index_name):
                df = df.reset_index()
                df = df.set_index(index_name).sort_index()
        else:
            df = df.set_index(index_name).sort_index()
        for col in columns:
            if col in df.columns:
                cut_df = df.reindex(columns=[col])
                cut_df = cut_df.rename(columns={col: 'value'})
                if cut_df['value'].notnull().all():
                    cut_df['value'] = cut_df['value'].astype(float)
                    cut_df['col'] = col
                    ret = ret.append(cut_df)
                else:
                    print('column', col, 'has null')
        ret = ret[(ret['value'] != 0) & (ret['value'].notnull())]
        ret = ret.reset_index()
        return ret

    @classmethod
    def drop_repeat_columns(cls, df: pd.DataFrame):
        '''
        找出重复的列名，选择第一个列
        '''
        df = df.copy()
        repeate = df.columns[df.columns.duplicated()]
        for r in repeate:
            # 获取当前重复列名r的所有列，组成新的df
            t_df = df.loc[:, r]
            # 新的df列重命名
            t_df.columns = range(t_df.shape[1])
            # 原df抛弃当前重复列名r
            df.drop(r, axis=1, inplace=True)
            # 取第一个重复列名作为原df的列名补充回去
            df[r] = t_df[0]
        return df

    @classmethod
    def find_col_name(cls, df: pd.DataFrame, findname: str):
        '''
        搜索实际列名。由于实际列名不固定会需要
        '''
        columns = df.columns.to_series()
        columns = columns[columns.notnull()]
        t_columns = columns[columns == findname]
        if t_columns.empty:
            columns = columns[columns.str.startswith(findname)]
        else:
            columns = t_columns
        if columns.empty:
            raise ValueError('not found column as {}'.format(findname))
        return columns[0]

    @classmethod
    def move_col(cls, df: pd.DataFrame, col_name: Union[str, list], pos: int):
        if isinstance(col_name, str):
            col = df[col_name]
            df = df.drop(col_name, axis=1)
            df.insert(pos, col_name, col)
            return df
        elif isinstance(col_name, list):
            col_name.reverse()
            for col in col_name:
                df = Others.move_col(df, col, pos)
        return df

    @staticmethod
    def add_order_index(df_dict:Dict[str,pd.DataFrame], col_name='project'):
        df_list=[pd.DataFrame()]
        df_list.clear()
        order=0
        for k,v in df_dict.items():
            v=v.copy()
            v[col_name] = "%02d"%order + k
            if v.index.name==None and v.index.names==[None]:
                v.set_index(col_name,inplace=True)
            else:
                v.set_index(col_name,append=True,inplace=True)
            df_list.append(v)
            order+=1
        return pd.concat(df_list,axis=0).sort_index()

    @staticmethod
    def clean_orderindex_mark(df: pd.DataFrame, col_name='project'):
        '''Clean the sort number info in df[col_name].

        Args:
            df (pd.DataFrame): [description]
            col_name (str, optional): [description]. Defaults to 'project'.

        Returns:
            [type]: [description]
        '''
        idx_list = df.index.names

        df = df.sort_index().reset_index()

        def clean(x: str):
            return x.strip('1234567890')

        df[col_name] = df[col_name].apply(clean)

        ret = df.set_index(idx_list)

        return ret

    @staticmethod
    def get_differ(series_A: pd.Series, series_B: pd.Series):
        '''
        得到两个Series差集:series_A比series_B多的数据
        1.通过添加两次相同的数据使需要排除的数据必定重复
        2.使用<去重——不保留>方法获得两个Series不重复的数据
        3.因为df2被添加两次，所以df2比df1多的数据不会保留下来
        '''

        series_A = series_A.drop_duplicates(keep='first')
        series_B = series_B.drop_duplicates(keep='first')
        series_A = series_A.append(series_B).append(series_B)
        diff = series_A.drop_duplicates(keep=False)
        diff.sort_values(inplace=True, ignore_index=True)
        return diff
