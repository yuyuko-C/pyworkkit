import smtplib
import socket
from .email import Email, parseaddr


class STMP_SSL_Sever(smtplib.SMTP_SSL):
    def __init__(self, host: str, port: int) -> None:
        super().__init__(host=host, port=port, local_hostname=None, keyfile=None, certfile=None, timeout=socket._GLOBAL_DEFAULT_TIMEOUT,
                         source_address=None, context=None)

    @staticmethod
    def get_to_addrs(item: Email):
        to_addrs = item.to_addrs
        carbons = item.carbons
        ret = []
        if to_addrs and isinstance(to_addrs, list):
            ret.extend([parseaddr(rec)[1] for rec in to_addrs])
        if carbons and isinstance(carbons, list):
            ret.extend([parseaddr(car)[1] for car in carbons])
        return ret

    def send_message(self, msg: Email, sender: str, to_addrs: list = None):
        if not msg.editable():
            msg.instance(force=False)

        msg.sender = sender

        to_addrs = self.get_to_addrs(msg) if to_addrs is None else to_addrs

        send_back = super().send_message(msg, self.user, to_addrs)
        return send_back
