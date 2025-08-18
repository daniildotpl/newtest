import pandas as pd
from sqlalchemy import text
from ishch.connector import my_connection



def get_payload(engine, query): 
    query = text(query)
    df = pd.read_sql(query, engine)
    df.to_excel('test_output.xlsx', index=False)
    print(df)

if __name__ == "__main__":
    get_payload(my_connection(), "SELECT * FROM statinfo;")

