

from easygui import fileopenbox, msgbox


import csv


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
def get_csv_file():
    fname = ''
    while len(fname) == 0:
        fname = fileopenbox("Оберіть файл CSV", default='*.csv')
        if not fname.endswith(".csv"):
            msgbox("Обрано не CSV- файл", ok_button="ОК", title="Перевірте тип файла!")
            fname = ''
            # exit ( )
    return fname


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
current_category = 0
checked_categories = []  # перелік категорій, які ми вже обробили

nonstandard_categories = ["І.ТОВАРИ", # перелік всіх категорій, які повинні бути у звіті
                          "ІІ.РОБОТИ",
                          "ІІІ.ПОСЛУГИ",
                          "Усього за розділом І (ТОВАРИ)",
                          "Усього за розділом ІІ (РОБОТИ)",
                          "Усього за розділом ІІІ (Послуги)",
                          "Разом (розділ І+розділ ІІ+розділ ІІІ)",
                          "КВЕД"
                         ]


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
# Функція зчитує СSV-файл
def read_from_csv(fname):
    global current_category
    csv.register_dialect("edata", delimiter=';')  # Символ - розділювач токенів

    try:
        with open(fname, 'rt', encoding='Windows-1251') as csv_file:
            row_count = 1
            for row in csv.DictReader(csv_file, dialect='edata'):
                row_count += 1

                for col in row:
                    # print (row[col])

                    validators[col](row[col], row_count)
    except FileNotFoundError:
        msgbox("Файл " + fname + " не знайдено!", ok_button="OK")


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


# =================== Функції - валідатори

def val_nomer_za_poryadkom(val, row_count=0):

    global current_category

    l = len(val)

    try:
        i = int(val)
    except ValueError:
        i = 0

    if l == 0: #Рядок порожній
        print('Рядок № %i : не вказано категорію (1,2 чи 3) або додаткові розділи!' % row_count)
        return
    
    if i > 0:
        if i != current_category:
            print('Рядок № %i : Вказана категорія %i не відповідає поточній %i !' % (row_count, i, current_category))
            return
    else:
        if (val in nonstandard_categories):
        # обробляємо випадок "Усього за розділом І (ТОВАРИ)" , "Усього за розділом ІІ (РОБОТИ)" ,
        # "Усього за розділом ІІІ (Послуги)" - стовпчик row[ '№з/п' ]
            checked_categories.append(val)

            if str(val).find("Усього за розділом") > 0:
                if len(row['Обсяг платежів, грн.']) == 0:
                    print('Рядок № %i : Відсутній підсумок за розділом %i!' % (row_count, n))
                    return
            elif str(val).find('Разом') > 0:
            # обробляємо випадок "Разом (розділ І+розділ ІІ+розділ ІІІ)" ,
                if len(row['Обсяг платежів, грн.']) == 0:
                    print('Рядок № %i : Відсутній загальний підсумок за усіма розділами!' % (row_count))
                    return
            elif str(val).find('КВЕД') > 0:
                if len(row['Обсяг платежів, грн.']) == 0:
                    print('Рядок № %i : Відсутній код КВЕД!' % (row_count))
                    return

        return

# =============================================================================
def val_vid(val, row_count=0):
    l = len(val)
    if l == 0:
        print('Рядок № %i : Відсутня дата договору!' % row_count)
        return
    elif 0 < l < 10:
        print('Рядок № %i : Невірно вказано дату договору!' % row_count)
        print('Дата повинна мати формат ДД.ММ.РРРР.')
        return
    return


def val_erdpou(val, row_count=0):
    l = len(val)
    if l == 0:
        print('Рядок № %i : Відсутній код ЕРДПОУ!' % row_count)
        return
    elif (l < 6) or (l > 10):
        print('Рядок № %i : Невірно вказаний код ЕРДПОУ! Код повинен мати довжину від 6 до 10 символів.' % row_count)
        return


def val_dogovir_no(val, row_count=0):
    if len(val) == 0:
        print('Рядок № %i : Відсутній номер договору!' % row_count)
        return
    elif len(val) > 30:
        print('Рядок № %i : довжина номеру договору перевищує 30 символів!' % row_count)
        return


def val_nazva_kontragenta(val, row_count=0):
    global current_category
    l = len(val)

    if val in nonstandard_categories:
        # обробляємо випадок "І.ТОВАРИ" , "ІІ.РОБОТИ" , "ІІІ.ПОСЛУГИ"
        checked_categories.append(val)
        current_category = nonstandard_categories.index(val) + 1
        return
    elif l == 0:
        print('Рядок № %i : Відсутня назва контрагента' % row_count)
        return
    elif (l > 120):
        print('Рядок № %i : Довжина назви контрагента перевищуе 120 символів.' % row_count)
        return



def val_predmet_dogovoru(val, row_count=0):
    l = len(val)
    s = str(val)
    if l == 0:
        print('Рядок № %i : Не заповнено предмет договору' % row_count)
    elif (l > 256):
        print('Рядок № %i : Довжина предмету договору перевищуе 256 символів.' % row_count)

    return


def val_kod_cpv_rozdil(val, row_count=0):
    l = len(val)
    s = str(val)
    if l == 0:
        print('Рядок № %i : Відсутній код СPV' % row_count)

    else:
        if s.find(' ') > 0 or s.find('"') > 0:
            print('Рядок № %i : Знайдені недопустимі символи-пробіл або лапки' % row_count)
        elif s.count('-') > 1:
            print('Рядок № %i : Мае бути тільки один символ прочерку' % row_count)
    return


def val_obsyag_plategiv_grn(val, row_count=0):
    l = len(val)
    if l == 0:
        print('Рядок № %i : Не вказано обсяг платежів!' % row_count)
    elif str(val).find(',') > 0:
        print('Рядок № %i : Копійки або код КВЕД в обсязи платежів повинні відокремлюватися '
              'тільки крапкою!' % row_count)
    return



def val_vidsutni_rozdili():
    return list(set(nonstandard_categories)-set(checked_categories))

# ======================================================
# встановлює зв'язок між стовпчиком та функцією для перевірки вмісту стовпчика

validators = {
    '№з/п': val_nomer_za_poryadkom,  # тут також перевіряємо проміжні та кінцеві підсумки
    'ЄДРПОУ контрагента': val_erdpou,
    'Договір №': val_dogovir_no,
    'від': val_vid,
    'Назва контрагента': val_nazva_kontragenta,  # тут також перевіряємо категорію
    'Предмет договору': val_predmet_dogovoru,
    'Код CPV (розділ)': val_kod_cpv_rozdil,
    'Обсяг платежів, грн.': val_obsyag_plategiv_grn
}

# =============================================================================================

fname = get_csv_file()
read_from_csv(fname)
print('Можливо, відсутні розділи: ',val_vidsutni_rozdili())
