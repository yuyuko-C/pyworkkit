from .path import Path
from .filesync import SyncGroup
from typing import Union



# 通过字符串特征搜索文件路径
# 返回值：结果唯一时为字符串，结果多个时为字符串列表，无结果时为Null


def find_file_paths(path, include: Union[list, str], exclude: Union[list, str] = None):

    if not isinstance(include, (list, str)):
        raise TypeError('include Type Error. got {}'.format(type(include)))
    if (exclude != None) and (not isinstance(exclude, (list, str))):
        raise TypeError('exclude Type Error. got {}'.format(type(exclude)))

    def core_code(include: list, exclude: list):
        pth = Path(path).instance()
        paths = ['']
        paths.clear()
        for file_name in pth.listdir():
            if any([word in file_name for word in include]):
                if (exclude != None) and any([word in file_name for word in exclude]):
                    continue
                paths.append(pth.join(file_name).path)
        return paths

    include = [include] if isinstance(include, str) else include
    exclude = [exclude] if isinstance(exclude, str) else exclude


    r_list = core_code(include, exclude)

    if len(r_list) == 0:
        return None
    elif len(r_list) == 1:
        return r_list[0]
    else:
        return r_list
