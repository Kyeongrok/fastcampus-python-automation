

import pandas as pd
from datetime import datetime, timedelta

start_date = '20220801'
dt = datetime.strptime('20220801', '%Y%m%d') + timedelta(days=6)
rct3 = pd.date_range(start=start_date, end=dt.strftime("%Y%m%d"))
dt_list = rct3.strftime("%Y-%m-%d").to_list()

print(dt_list)