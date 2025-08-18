import os
from sqlalchemy import create_engine
from dotenv import load_dotenv


load_dotenv()

dialect = os.getenv('dialect')
driver = os.getenv('driver')
dbhost = os.getenv('dbhost')
dbport = os.getenv('dbport')
dbname = os.getenv('dbname')
dbuser = os.getenv('dbuser')
dbpass = os.getenv('dbpass')
dbdriv = os.getenv('dbdriv')
# print(f'{dialect}:{driver}:{dbuser}:{dbpass}:{dbhost}:{dbport}:{dbname}')


def my_connection():
    connection_string = (
        f"{dialect}+{driver}://{dbuser}:{dbpass}@{dbhost}:{dbport}/{dbname}?{dbdriv}"
    )
    try:
        engine = create_engine(connection_string)
        with engine.connect() as conn:
            print("Successful connection")
            return engine
    except Exception as e:
        print(f"Error: {e}")
        return False


def str_to_float(str):
    if str.isdigit():
        return float(str)
    else:
        fl = str.replace(',', '.')
        return float(fl)


def add_cell_to_payload(payload, key):
    try:
        # value = round((str_to_float(payload.loc[1][key]) - str_to_float(payload.loc[0][key]))/str_to_float(payload.loc[1][key]), 2)
        value = round(((payload.loc[1][key] - payload.loc[0][key])/payload.loc[1][key]), 2)
    except:
        value = 0
    # print(value)
    return value