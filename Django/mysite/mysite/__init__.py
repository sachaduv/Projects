import os
import pymysql

pymysql.install_as_MySQLdb()
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))