import pandas as pd
from ishch.connector import my_connection, add_cell_to_payload
from docxtpl import DocxTemplate


def get_payload(engine):

    query = f"""
        SELECT 
            year
            , region
            , SUM(CASE 
                WHEN name_value = 'Число должностей врачей, ед занятых' 
                THEN TRY_CAST(REPLACE(REPLACE(value, ' ', ''), ',', '.') AS DECIMAL(15,2)) 
                ELSE 0 END) AS doctors
            , SUM(CASE 
                WHEN name_value = 'Число посещений к врачам, включая профилактические (без посещений к стоматологам и зубным врачам), ед' 
                THEN TRY_CAST(REPLACE(REPLACE(value, ' ', ''), ',', '.') AS DECIMAL(15,2)) 
                ELSE 0 END) AS doctors_visits
            , SUM(CASE 
                WHEN name_value = 'Число должностей среднего медперсонала, ед занятых' 
                THEN TRY_CAST(REPLACE(REPLACE(value, ' ', ''), ',', '.') AS DECIMAL(15,2)) 
                ELSE 0 END) AS nurses
            , SUM(CASE 
                WHEN name_value = 'Число посещений к среднему медперсоналу, ед' 
                THEN TRY_CAST(REPLACE(REPLACE(value, ' ', ''), ',', '.') AS DECIMAL(15,2)) 
                ELSE 0 END) AS nurses_visits
            , SUM(CASE 
                WHEN name_value = 'Число организаций, ед' 
                THEN TRY_CAST(REPLACE(REPLACE(value, ' ', ''), ',', '.') AS DECIMAL(15,2)) 
                ELSE 0 END) AS count_org
        FROM statinfo
        WHERE year IN (2021, 2022) 
        AND region = 2290
        GROUP BY year, region;
        ;
        """
    
    payload = pd.read_sql(query, engine)
    # print(str_to_float(payload.loc[1]['doctors']))


    payload.loc[len(payload)] = [
        'delta', 
        '-',
        add_cell_to_payload(payload, 'doctors'),
        add_cell_to_payload(payload, 'doctors_visits'),
        add_cell_to_payload(payload, 'nurses'),
        add_cell_to_payload(payload, 'nurses_visits'),
        add_cell_to_payload(payload, 'count_org')
    ]
    print(payload)
    return payload


def get_context(payload):
    # count of more and less deltas ------------
    count_less = 0
    count_more = 0
    deltas = payload.loc[2].tolist()
    del deltas[0:2]

    for d in deltas:
        # print(d)
        if d < 0:
            count_less += 1
        if d > 0:
            count_more += 1
    # print(f'count_less: {count_less}, count_more: {count_more}')

    # # make context ------------------------------
    context = {}
    payload_as_dict = payload.set_index('year').to_dict('index')
    for year, dic in payload_as_dict.items():
        for key, value in dic.items():
            key = key + '_' + str(year)
            # print(key)
            context[key] = str(value)
    # # ---
    context['count_less'] = str(count_less)
    context['count_more'] = str(count_more)
    print(f'context: {context}')
    return context


def make_word(context):
    # make dokument -----------------------------
    doc = DocxTemplate("ishch/template.docx")
    try:
        doc.render(context)
    except Exception as error:
        print(f'Ошибка в render: {error}')

    doc.save("generated_document.docx")


if __name__ == "__main__":
    payload = get_payload(my_connection())
    context = get_context(payload)
    make_word(context)













        # WITH preset AS (
        #     SELECT 
        #         year
        #         , region
        #         , name_value
        #         , value 
        #     FROM statinfo
        #     WHERE year IN (2021, 2022) 
        #     AND region=2290
        # )
        # SELECT 
        #     year
        #     , region
            
        #     , (SELECT TOP 1 value 
        #         FROM statinfo 
        #             WHERE name_value='Число должностей врачей, ед занятых'
        #             AND year=preset.year
        #             AND region=preset.region
        #             AND tip_mo=tip_mo) as doctors

        #     , (SELECT TOP 1 value 
        #         FROM statinfo 
        #             WHERE name_value='Число посещений к врачам, включая профилактические (без посещений к стоматологам и зубным врачам), ед'
        #             AND year=preset.year
        #             AND region=preset.region
        #             AND tip_mo=tip_mo) as doctors_visits

        #     , (SELECT TOP 1 value 
        #         FROM statinfo 
        #             WHERE name_value='Число должностей среднего медперсонала, ед занятых'
        #             AND year=preset.year
        #             AND region=preset.region
        #             AND tip_mo=tip_mo) as nurses

        #     , (SELECT TOP 1 value 
        #         FROM statinfo 
        #             WHERE name_value='Число посещений к среднему медперсоналу, ед'
        #             AND year=preset.year
        #             AND region=preset.region
        #             AND tip_mo=tip_mo) as nurses_visits

        #     , (SELECT TOP 1 value 
        #         FROM statinfo 
        #             WHERE name_value='Число организаций, ед'
        #             AND year=preset.year
        #             AND region=preset.region
        #             AND tip_mo=tip_mo) as count_org

        # FROM preset
        # GROUP BY year, region
        # ;