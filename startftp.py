from pyftpdlib.authorizers import DummyAuthorizer
from pyftpdlib.handlers import FTPHandler, ThrottledDTPHandler
from pyftpdlib.servers import FTPServer
from pyftpdlib.log import LogFormatter
import logging

logger = logging.getLogger('FTP-LOG')
logger.setLevel(logging.DEBUG)

cs = logging.StreamHandler()
cs.setLevel(logging.INFO)

fs = logging.FileHandler(filename='test.log', mode='a', encoding='utf-8')
fs.setLevel(logging.DEBUG)

formatter = logging.Formatter('[%(asctime)s] %(name)s - %(levelname)s : %(message)s')

cs.setFormatter(formatter)
fs.setFormatter(formatter)

logger.addHandler(cs)
logger.addHandler(fs)

auth = DummyAuthorizer()

auth.add_user('user', '123123', "e:/sushun", perm='elradfmw')

handler = FTPHandler
handler.authorizer = auth
handler.passive_ports = range(2000, 20033)

dtp_handler = ThrottledDTPHandler
dtp_handler.read_limit = 3000 * 1024
dtp_handler.write_limit = 3000 * 1024

handler.dtp_handler = dtp_handler

server = FTPServer(('0.0.0.0', 21), handler)
server.max_cons = 100
server.max_cons_per_ip = 10

server.serve_forever()