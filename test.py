from datetime import datetime

date_str = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
date = date_str.replace(':', '')
print(date)