import pandas as pd
from ishch.connector import my_connection
from sqlalchemy import text



def get_payload(engine):
    query = f"""
        WITH preset AS (
            SELECT 
                year
                , region
                , tip_mo
                , name_value
                , value visits 
                , 100*COUNT(*) OVER (PARTITION BY region, tip_mo)/COUNT(*) OVER (PARTITION BY region) peroftype
            FROM statinfo
            WHERE 
            name_value='Число посещений к врачам, включая профилактические (без посещений к стоматологам и зубным врачам), ед' 
        )
        SELECT 
            year AS [YEAR]
            , region AS [REGION]
            , tip_mo AS [TYPE]
            , peroftype AS [PERC. OF TIP MO]
            , DENSE_RANK() OVER (PARTITION BY region ORDER BY peroftype DESC) AS [RANK PEROFTYPE]
            , visits AS [VISITS]
            , ROUND(TRY_CONVERT(FLOAT, visits)/(SELECT TOP 1 population 
                        FROM population 
                        WHERE region=region AND year=year)*100000, 1) AS [VISITS ON TH]

        FROM preset 
        ORDER BY year
    ;
    """
    df = pd.read_sql(query, engine)
    return df


if __name__ == "__main__":
    print(get_payload(my_connection()))

