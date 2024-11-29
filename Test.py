from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

# 获取今天的日期
today = datetime.today()

# 获取下个月的第一天
first_day_of_next_month = (today + relativedelta(months=+1)).replace(day=1)

# 获取本月最后一天
last_day_of_month = first_day_of_next_month - timedelta(days=1)

# 计算今天到月底的天数
days_until_end_of_month = (last_day_of_month - today).days

print(f"今天到月底还有 {days_until_end_of_month} 天")