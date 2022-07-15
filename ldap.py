import ldap3
import os
from ldap3.core.exceptions import LDAPBindError




class GetNameFromLdap(object):
    def __init__(self, server, user, password):
        self.server = ldap3.Server(server, get_info=ldap3.ALL)
        self.user = user
        self.password = password


    def Cnname(self):
        try:
            self.conn = ldap3.Connection(self.server, user=self.user, password=self.password, auto_bind=True,
                                         authentication=ldap3.NTLM)
            if self.conn.bind():
                AD_SEARCH_TREE = 'dc=corp,dc=primatek,dc=ru'

                display_name = str(os.getlogin())
                self.conn.search(AD_SEARCH_TREE, f'(&(objectCategory=Person)(sAMAccountName={display_name}))',
                                 attributes=['cn'])
                result = str(self.conn.entries[0])
                index_cn = result.find('cn:')
                rez = result[(index_cn + 4):]
                return str(rez)
        except LDAPBindError as s:
            print(f"Введены не верные учетные данные, {s}")

