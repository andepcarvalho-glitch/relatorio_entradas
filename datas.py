from datetime import datetime, timedelta

class Datas:
    @staticmethod
    def get():
        hoje = datetime.now()
        inicio_mes = hoje.replace(month=hoje.month-1, day=1).strftime("%Y-%m-%d")
        if hoje.month == 12:
            fim_mes = hoje.replace(month=12, day=31)
        else:
            fim_mes = (hoje.replace(month=hoje.month+1, day=1) - timedelta(days=1)).strftime("%Y-%m-%d")
        return inicio_mes, fim_mes
