import datetime
import mysql.connector
import requests
from mysql.connector import Error
import pandas as pd
import matplotlib.pyplot as plt

def add_btc():
  try:
    mydb = mysql.connector.connect(
      host="localhost",
      user="root",
      password="verga",
      database="sahaquiel"
    )
    if mydb.is_connected():
        
      cursor = mydb.cursor()

      cursor.execute('''CREATE TABLE IF NOT EXISTS btc
                    (id INT AUTO_INCREMENT PRIMARY KEY,
                    date DATETIME NOT NULL,
                    price FLOAT NOT NULL);''')
      date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

      response = requests.get('https://api.coinbase.com/v2/prices/BTC-USD/spot')

      btc_price = response.json()['data']['amount']

      query = "INSERT INTO btc (date, price) VALUES (%s, %s)"
      values = (date, btc_price)
      cursor.execute(query, values)
      mydb.commit()

      cursor = mydb.cursor()
      cursor.execute("SELECT * FROM btc")
      rows = cursor.fetchall()

      df = pd.DataFrame(rows, columns=['id', 'date', 'price'])
      df.to_excel('btc.xlsx', index=False)

      df = pd.read_excel("btc.xlsx")
      daily_prices = df.groupby('date')['price'].mean()
      daily_prices.plot(kind='bar')
      plt.title('BTC Price')
      plt.xlabel('Date')
      plt.ylabel('Price')
      plt.show()
       
  except Error as e:
      print('Error:', e)

  finally:
      if mydb.is_connected():
          cursor.close()
          mydb.close()

def main():
  add_btc()


if __name__ == "__main__":
    main()