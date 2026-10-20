# Flask - app.py
import os
import requests
import pandas as pd

url = "https://api.pluggy.ai"

def authenticate():
    print(os.getenv("CLIENT_ID"))
    print(os.getenv("CLIENT_SECRET"))
    payload = {
        "clientId": str(os.getenv("CLIENT_ID")),
        "clientSecret": str(os.getenv("CLIENT_SECRET"))
    }
    headers = {
        "accept": "application/json",
        "content-type": "application/json"
    }

    return requests.post(url+"/auth", json=payload, headers=headers)

auth_response = authenticate().json().get("apiKey")
print(auth_response)

headers = {
    "X-API-KEY": auth_response,
    "accept": "application/json"
}

response = requests.get(url+"/accounts?itemId=" + str(os.getenv("ITEM_ID")), headers=headers).json()
accounts = []

for account in response.get("results"):

    cleanedTransactions = []
    transactionPath = "/v2/transactions"
    accountId = str(account.get("id"))
    accountIdPath = "?accountId=" + accountId

    has_more = True
    print("Retrieving transactions for account ID:", accountId)

    while has_more:
        response = requests.get(url+transactionPath + accountIdPath, headers=headers)
        
        if response.status_code != 200:
            print(f"Request Error: {response.status_code} - {response.text}")
            break

        data = response.json()
        transactions = data.get("results", [])
        
        for transaction in transactions:
            cleanTransaction = {
                "id": transaction.get("id"),
                "description": transaction.get("description"),
                "currencyCode": transaction.get("currencyCode"),
                "amount": transaction.get("amount"),
                "date": transaction.get("date"),
                "balance": transaction.get("balance"),
                "category": transaction.get("category"),
                "accountId": transaction.get("accountId"),
                "status": transaction.get("status"),
                "type": transaction.get("type"),
                "accountNumber": account.get("number")
            }
            cleanedTransactions.append(cleanTransaction)

        next_cursor = data.get("next")
        print(data)
        if next_cursor:
            print("Next cursor found")
            accountIdPath = next_cursor
            print(accountIdPath)
        else:
            has_more = False

    sortedTransactions = sorted(cleanedTransactions, key=lambda x: x['date'])

    if os.path.exists("Bionte\Trabalho\OpenFinance\livro_razao.xlsx"):
        with pd.ExcelWriter(os.path.join('Bionte\Trabalho\OpenFinance', "livro_razao.xlsx"), engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
            dataframe = pd.DataFrame(sortedTransactions)
            dataframe.to_excel(writer, sheet_name="razão-"+account.get("id"), index=False)

    else:
        with pd.ExcelWriter(os.path.join('Bionte\Trabalho\OpenFinance', "livro_razao.xlsx"), engine="openpyxl") as writer:
            dataframe = pd.DataFrame(sortedTransactions)
            dataframe.to_excel(writer, sheet_name="razão-"+account.get("id"), index=False)

    print("Transactions retrieval completed for account ID:", accountId)