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
                , SUM(CASE 
                    WHEN name_value = 'Число организаций, ед' 
                    THEN TRY_CAST(REPLACE(REPLACE(value, ' ', ''), ',', '.') AS DECIMAL(15,2)) 
                    ELSE 0 END) AS count_org
                , SUM(CASE 
                    WHEN name_value = 'Число посещений к врачам, включая профилактические (без посещений к стоматологам и зубным врачам), ед' 
                    THEN TRY_CAST(REPLACE(REPLACE(value, ' ', ''), ',', '.') AS DECIMAL(15,2)) 
                    ELSE 0 END)/(SELECT TOP 1 population 
                        FROM population 
                        WHERE region=region AND year=year)*100000 AS doctors_visits_on_th
            FROM statinfo
            GROUP BY year, region, tip_mo 
        )
        SELECT 
            year
            , region
            , tip_mo
            , count_org
            , DENSE_RANK() OVER (PARTITION BY region ORDER BY count_org DESC) AS count_org_rank
            , doctors_visits_on_th
            , DENSE_RANK() OVER (PARTITION BY region ORDER BY doctors_visits_on_th DESC) AS doctors_visits_on_th_rank
        FROM preset
    ;
    """
    df = pd.read_sql(query, engine)
    return df


if __name__ == "__main__":
    print(get_payload(my_connection()))

