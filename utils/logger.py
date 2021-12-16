#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Author: zhangjun
# @Date  : 2018/7/26 9:20
# @Desc  : Description
 
import logging
import logging.handlers
import os
import time
import typing
 
class Logs(object):
    __loggerDic:typing.Dict[str,logging.Logger]={}

    @classmethod
    def getLogger(cls,name:str):
        return cls.__loggerDic[name]

    @classmethod
    def createLogger(cls,name:str,logs_dir:str,level:str):
        if name not in cls.__loggerDic.keys():
            cls.__setup__(name,logs_dir,level)

    @classmethod
    def __setup__(cls,name:str,logs_dir:str,level:str):
        logger = logging.getLogger(name)
        cls.__loggerDic[name]=logger
        # 设置输出的等级
        LEVELS = {'NOSET': logging.NOTSET,
                  'DEBUG': logging.DEBUG,
                  'INFO': logging.INFO,
                  'WARNING': logging.WARNING,
                  'ERROR': logging.ERROR,
                  'CRITICAL': logging.CRITICAL}
        level = LEVELS[level.upper()]
        # 创建文件目录
        if os.path.exists(logs_dir) and os.path.isdir(logs_dir):
            pass
        else:
            os.mkdir(logs_dir)
        # 修改log保存位置
        timestamp=time.strftime("%Y-%m-%d",time.localtime())
        logfilename='{0}-{1}.txt'.format(name,timestamp)
        logfilepath=os.path.join(logs_dir,logfilename)
        rotatingFileHandler = logging.handlers.RotatingFileHandler(filename =logfilepath,
                                                                   maxBytes = 1024 * 1024 * 50,
                                                                   backupCount = 5)
        # 设置输出格式
        formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', '%Y-%m-%d %H:%M:%S')
        rotatingFileHandler.setFormatter(formatter)
        # 控制台句柄
        console = logging.StreamHandler()
        console.setLevel(level)
        console.setFormatter(formatter)
        # 添加内容到日志句柄中
        logger.addHandler(rotatingFileHandler)
        logger.addHandler(console)
        logger.setLevel(level)

