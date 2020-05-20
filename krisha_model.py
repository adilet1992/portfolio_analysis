import pandas as pd

def predictPriceModel(model, year, rooms, Fl, Max, Squares, types, district):
    type_кирпичный = 0
    type_монолитный = 0
    type_панельный = 0
    District_Алатауский = 0
    District_Алмалинский = 0
    District_Ауэзовский = 0
    District_Бостандыкский = 0
    District_Жетысуский = 0
    District_Медеуский = 0
    District_Наурызбайский = 0
    District_Турксибский = 0
    if (types == 'кирпичный'):
        type_кирпичный = 1
    if (types == 'монолитный'):
        type_монолитный = 1
    if (types == 'панельный'):
        type_панельный = 1
    if (district == 'Алатауский'):
        District_Алатауский = 1
    if (district == 'Алмалинский'):
        District_Алмалинский = 1
    if (district == 'Ауэзовский'):
        District_Ауэзовский = 1
    if (district == 'Бостандыкский'):
        District_Бостандыкский = 1
    if (district == 'Жетысуский'):
        District_Жетысуский = 1
    if (district == 'Медеуский'):
        District_Медеуский = 1
    if (district == 'Наурызбайский'):
        District_Наурызбайский = 1
    if (district == 'Турксибский'):
        District_Турксибский = 1
    inputs = []
    inputs.append({'year':year, 'rooms':rooms, 'Fl':Fl, 'Max':Max, 'Squares':Squares, 'type_кирпичный':type_кирпичный,
                    'type_монолитный':type_монолитный, 'type_панельный':type_панельный, 'District_Алатауский':District_Алатауский,
                    'District_Алмалинский':District_Алмалинский, 'District_Ауэзовский':District_Ауэзовский,
                    'District_Бостандыкский':District_Бостандыкский, 'District_Жетысуский':District_Жетысуский,
                    'District_Медеуский':District_Медеуский, 'District_Наурызбайский':District_Наурызбайский,
                    'District_Турксибский':District_Турксибский})
    inp = pd.DataFrame(inputs)
    inp = inp.reindex(['year', 'rooms', 'Fl', 'Max', 'Squares', 'type_кирпичный', 'type_монолитный', 'type_панельный', 'District_Алатауский', 'District_Алмалинский', 'District_Ауэзовский', 'District_Бостандыкский', 'District_Жетысуский', 'District_Медеуский', 'District_Наурызбайский', 'District_Турксибский'], axis=1)
    y3 = model.predict(inp)
    
    return y3[0]