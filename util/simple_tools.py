from datetime import datetime

class SimpleTools:
    def dete_format(self, date):
        """
        格式化如2023-02-15T10:13:29.001+0800的日期格式
        :param date:
        :return:
        """
        t = date[:-9]
        local_date = datetime.strptime(t, "%Y-%m-%dT%H:%M:%S")
        return local_date
