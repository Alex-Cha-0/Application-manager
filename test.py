import os
from datetime import datetime
import time

import pymssql
import pytz
from exchangelib import DELEGATE, Account, Credentials, Configuration
import exchangelib.autodiscover
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from requests_kerberos import HTTPKerberosAuth
import urllib3


credentials = Credentials('helper@primatek.ru', 'Renovation21')

exchangelib.transport.AUTH_TYPE_MAP["GSSAPI"] = auth_model

config = Configuration(server='mimir.corp.primatek.ru', credentials=credentials, auth_type="GSSAPI")

account = Account(primary_smtp_address='helper@primatek.ru', credentials=credentials, autodiscover=False,config=config, access_type=DELEGATE)


# Use this adapter class instead of the default
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

for item in account.inbox.all().order_by('-datetime_received')[:1]:
    print(item.subject, item.sender, item.datetime_received)