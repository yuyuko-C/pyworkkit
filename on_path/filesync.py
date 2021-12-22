from .path import Path
import typing


class SyncGroup:

    def __init__(self, source: str, target: str, *filters: str, filter_mode: bool = True) -> None:
        source_pth = Path(source).absolute()
        target_pth = Path(target).absolute()
        if source_pth.is_dir() and target_pth.is_dir():
            self.source = source_pth
            self.target = target_pth
        else:
            raise ValueError("")

        self.__set_filter(*filters, filter_mode=filter_mode)

    def __set_filter(self, *filters: str, filter_mode: bool):
        filters_path = []
        for p_filter in filters:
            real_f = self.source.joinpath(p_filter).relative_to(self.source)
            filters_path.append(real_f)
        self.__path_filters:typing.List[Path] = filters_path
        self.__filters_mode:bool = filter_mode

    def __filter_check(self, relative_path: Path):
        for p_filter in self.__path_filters:
            if relative_path.is_relative_to(p_filter):
                return not self.__filters_mode
        return self.__filters_mode

    def __filter_parent_check(self, relative_path: Path):
        for p_filter in self.__path_filters:
            if p_filter.is_relative_to(relative_path):
                return not self.__filters_mode
        return self.__filters_mode

    def __clean_target(self):
        # 先比较文件夹
        for target in self.target.walk_dir():
            relpath = target.relative_to(self.target)
            # 检查此路径能否通过过滤器
            if self.__filter_check(relpath) or self.__filter_parent_check(relpath):
                source = self.source.joinpath(relpath)
                # 如果源路径被删除，也要移除
                if not source.exists():
                    target.rmdir(True)
            else:
                # 不能通过过滤器则不该同步，直接移除
                target.rmdir(True)


        # 再比较文件
        for target in self.target.walk_file():
            relpath = target.relative_to(self.target)
            # 检查此路径能否通过过滤器
            if self.__filter_check(relpath):
                source = self.source.joinpath(relpath)
                # 如果源路径被删除，也要移除
                if not source.exists():
                    target.unlink()
            else:
                # 不能通过过滤器则不该同步，直接移除
                target.unlink()

    def __source_to_target(self):
        # 遍历没有的文件夹，复制并收录到copy_folder
        copy_folder = []
        for source in self.source.walk_dir():
            relpath = source.relative_to(self.source)
            # 检查此路径能否通过过滤器
            if self.__filter_check(relpath):
                target = self.target.joinpath(relpath)
                if not target.exists():
                    source.copy_to(target)
                    copy_folder.append(source)

        # 遍历文件，跳过copy_folder中的文件夹，若文件修改时间不同则判定为不同
        for source in self.source.walk_file(*copy_folder):
            relpath = source.relative_to(self.source)
            # 检查此路径能否通过过滤器
            if self.__filter_check(relpath):
                target = self.target.joinpath(relpath)
                if not target.exists():
                    source.copy_to(target)
                else:
                    if target.stat().st_mtime != source.stat().st_mtime:
                        source.copy_to(target)

    def execute_same(self):
        # 1.检查target比source多出来的路径，删除
        # 2.检查source比target多出来的路径，复制
        # 3.检查source与target一致的路径，检查
        # 1.如果文件信息一致，不变
        # 2.如果文件信息不一致，替换

        self.__clean_target()
        self.__source_to_target()
        pass
    
    def execute_addition(self):
        # 1.检查source比target多出来的路径，复制
        # 2.检查source与target一致的路径，检查
        # 1.如果文件信息一致，不变
        # 2.如果文件信息不一致，替换

        self.__source_to_target()
        pass
