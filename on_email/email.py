from pathlib import Path
from typing import Union
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.utils import parseaddr, formataddr

from on_email.error import OnceError
from on_email.html import Html_Maker


def before_instance(method):
    def inner(self, *args, **kwargs):
        if self.editable():
            method(self, *args, **kwargs)
        else:
            raise AttributeError("instance variable can not be assigned.")
    return inner


def var_to_list(s):
    if not s:
        return []
    if isinstance(s, list):
        return s
    elif isinstance(s, tuple):
        return list(s)
    else:
        return [s]


class Email(MIMEMultipart):
    def __init__(self, title: str) -> None:
        super().__init__(_subtype='mixed', boundary=None, _subparts=None, policy=None)

        self.__instance = False
        self.__encode = 'utf-8'
        self.__files = {}
        self._source = {'title': '', 'sender': '',
                        'text': '', 'toaddrs': [], 'carbons': []}

        self.title = title

    @property
    def title(self):
        return self._source['title']

    @title.setter
    @before_instance
    def title(self, value: str):
        if not self._source['title']:
            self['Subject'] = Header(value, self.__encode).encode()
            self._source['title'] = value
        else:
            raise OnceError('title')

    @property
    def sender(self):
        return self._source['sender']

    @sender.setter
    @before_instance
    def sender(self, value: str):
        if not self._source['sender']:
            self['From'] = self.__format_addr(value)
            self._source['sender'] = value
        else:
            raise OnceError('sender')

    @property
    def text(self):
        return self._source['text']

    @text.setter
    @before_instance
    def text(self, value: str):
        self._source['text'] = value

    @property
    def to_addrs(self):
        return self._source['toaddrs']

    @to_addrs.setter
    @before_instance
    def to_addrs(self, value: Union[list, str]):
        if not self._source['toaddrs']:
            self._source['toaddrs'] = var_to_list(value)
            self['To'] = ','.join([self.__format_addr(rec)
                                  for rec in self._source['toaddrs']])
        else:
            raise OnceError('to_addrs')

    @property
    def carbons(self):
        return self._source['carbons']

    @carbons.setter
    @before_instance
    def carbons(self, value: Union[list, str]):
        if not self._source['carbons']:
            self._source['carbons'] = var_to_list(value)
            self['Cc'] = ','.join([self.__format_addr(car)
                                  for car in self._source['carbons']])
        else:
            raise OnceError('carbons')

    def add_attach(self, path: Path, type="attach"):
        if (path.exists() and path.is_file()):
            self.__files[path] = True
            file_id = len(self.__files)

            if type == "attach":
                with path.open('rb') as f:
                    file = MIMEApplication(f.read())
                    file.add_header('content-disposition',
                                    'attachment', filename=path.name)
                    self.attach(file)

            elif type == "image":
                with open(path, 'rb') as f:
                    mimage = MIMEImage(f.read())
                    mimage.add_header('Content-ID', str(file_id))
                    self.add_attach(mimage)

            return file_id
        else:
            raise FileExistsError(
                'not exists or is not file. check your path.')

    def insert_image(self, path: str, width: int = None, height: int = None):
        file_id = self.add_attach(path, type="image")
        image_html = Html_Maker.image(f'cid:{file_id}', width, height)
        self.text += image_html

    def struct_check(self):
        if not self.title:
            raise ValueError('e-mail title can not be empty')
        elif not self.text:
            raise ValueError('e-mail text can not be empty')
        elif not self.to_addrs:
            raise ValueError('e-mail to_addr can not be empty')

    def editable(self):
        return not self.__instance

    def instance(self, force: bool):
        if not force:
            self.struct_check()

        self.attach(MIMEText(self.text, 'html', self.__encode))
        self.__instance = True

    def __format_addr(self, s):
        name, addr = parseaddr(s)
        return formataddr((Header(name, self.__encode).encode(), addr))
