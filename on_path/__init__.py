from .path import Path
from .filesync import SyncGroup
import typing



# 通过字符串特征搜索文件路径
# 返回值：结果唯一时为字符串，结果多个时为字符串列表，无结果时为Null


def find_file_paths(path, include: typing.Union[list, str], exclude: typing.Union[list, str] = None):

    if not isinstance(include, (list, str)):
        raise TypeError('include Type Error. got {}'.format(type(include)))
    if (exclude != None) and (not isinstance(exclude, (list, str))):
        raise TypeError('exclude Type Error. got {}'.format(type(exclude)))
    include = [include] if isinstance(include, str) else include
    exclude = [exclude] if isinstance(exclude, str) else exclude

    pth = Path(path)
    paths:typing.List[str] = []
    for file_name in pth.iterdir():
        file_name = str(file_name)
        if any([word in file_name for word in include]):
            if file_name.startswith("~$") or (exclude and any([word in file_name for word in exclude])):
                continue
            paths.append(str(pth.joinpath(file_name)))

    if len(paths) == 0:
        return None
    elif len(paths) == 1:
        return paths[0]
    else:
        return paths
