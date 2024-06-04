import win32com.client
import pandas as pd
from datetime import datetime

class OutlookReport:
    def __init__(self):
        _outlook = win32com.client.Dispatch("Outlook.Application")
        self._namespace = _outlook.GetNamespace("MAPI")

    def get_messages(self):
        inbox = self._namespace.GetDefaultFolder(6)
        self._messages = inbox.Items

    def generate_report(self):
        self.get_messages()
        email_data = []
        for i in range(20):
            message = self._messages[i]
            try:
                email_data.append({
                    "Assunto": message.Subject,
                    "Data de Recebimento": message.ReceivedTime
                })
            except Exception as e:
                print(f"Erro ao processar o e-mail {e}")

        df = pd.DataFrame(email_data)
        df.to_csv("relatorio_emails.csv", index=False)
        print("Relatorio gerado com sucesso")

if __name__ == '__main__':
    report = OutlookReport()
    report.generate_report()