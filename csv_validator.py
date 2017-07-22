
# Увага! Для роботи програми потрібно встановити бібліотеку EasyGui
# pip install easygui

from easygui import fileopenbox , msgbox

import csv


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
def get_csv_file ():
    fname = ''
    while len ( fname ) == 0:
        fname = fileopenbox ( "Оберіть файл CSV" , default='*.csv' )
        if not fname.endswith ( ".csv" ):
            msgbox ( "Обрано не CSV- файл" , ok_button="ОК" , title="Перевірте тип файла!" )
            fname = ''
            # exit ( )
    return fname


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
current_category = 0
checked_categories = [ ]  # перелік категорій, які ми вже обробили

nonstandard_categories = [ "І.ТОВАРИ" ,  # перелік всіх категорій, які повинні бути у звіті
                           "ІІ.РОБОТИ" ,
                           "ІІІ.ПОСЛУГИ" ,
                           "Усього за розділом І (ТОВАРИ)" ,
                           "Усього за розділом ІІ (РОБОТИ)" ,
                           "Усього за розділом ІІІ (Послуги)" ,
                           "Разом (розділ І+розділ ІІ+розділ ІІІ)" ,
                           "КВЕД"
                           ]


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
# Функція зчитує СSV-файл
def read_from_csv ( fname ):
    global current_category
    csv.register_dialect ( "edata" , delimiter=';' )  # Символ - розділювач токенів
    prot = open ( "Протокол перевірки.txt" , "wt" )
    msg = ''
    try:
        with open ( fname , 'rt' , encoding='Windows-1251' ) as csv_file:
            row_count = 1
            for row in csv.DictReader ( csv_file , dialect='edata' ):
                row_count += 1

                for col in row:
                    # print (row[col])

                    msg = validators[ col ] ( row[ col ] , row_count , row )
                    if len(msg) == 0:
                        continue
                    prot.write (msg + '\r\n')
                # prot.write('\r\n')
        prot.write (val_vidsutni_rozdili() )
        prot.close ( )
    except FileNotFoundError:
        msgbox ( "Файл " + fname + " не знайдено!" , ok_button="OK" )


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


# =================== Функції - валідатори

def val_nomer_za_poryadkom ( val , row_count , row ):
    global current_category

    ryadok = 'Рядок № ' + str ( row_count ) + ':'
    msg = ''
    l = len ( val )
    try:
        i = int ( val )
    except ValueError:
        i = 0

    if l == 0:  # Рядок порожній

        msg += ryadok + ' не вказано категорію (1,2 чи 3) або додаткові розділи!'

    if i > 0:
        if i != current_category:
            msg += ryadok + ' Вказана категорія' + str ( i ) + ' не відповідає поточній' + str (
                current_category ) + '!'
    else:
        if (val in nonstandard_categories):
            # обробляємо випадок "Усього за розділом І (ТОВАРИ)" , "Усього за розділом ІІ (РОБОТИ)" ,
            # "Усього за розділом ІІІ (Послуги)" - стовпчик row[ '№з/п' ]
            checked_categories.append ( val )

            if str ( val ).find ( "Усього за розділом" ) > 0:
                if len ( row[ 'Обсяг платежів, грн.' ] ) == 0:
                    msg += ryadok + ' Відсутній підсумок за розділом' + val + '!'
            elif str ( val ).find ( 'Разом' ) > 0:
                # обробляємо випадок "Разом (розділ І+розділ ІІ+розділ ІІІ)" ,
                if len ( row[ 'Обсяг платежів, грн.' ] ) == 0:
                    msg += ryadok + ' Відсутній загальний підсумок за усіма розділами!'
            elif str ( val ).find ( 'КВЕД' ) > 0:
                if len ( row[ 'Обсяг платежів, грн.' ] ) == 0:
                    msg += ryadok + 'Відсутній код КВЕД!'

    return msg


# =============================================================================
def val_vid ( val , row_count , row ):
    l = len ( val )
    ryadok = 'Рядок № ' + str ( row_count ) + ':'
    msg = ''
    if l == 0:
        msg += ryadok + 'Відсутня дата договору!'
    elif 0 < l < 10:
        msg += ryadok + 'Дата договору повинна мати формат ДД.ММ.РРРР.'

    return msg


def val_erdpou ( val , row_count , row ):
    l = len ( val )
    ryadok = 'Рядок № ' + str ( row_count ) + ':'
    msg = ''
    if l == 0:
        msg += ryadok+'Відсутній код ЄРДПОУ!'
    elif (l < 6) or (l > 10):
        msg += ryadok+'Код ЄРДПОУ повинен мати довжину від 6 до 10 символів.'

    return msg


def val_dogovir_no ( val , row_count , row ):
    ryadok = 'Рядок № ' + str ( row_count ) + ':'
    msg = ''
    if len ( val ) == 0:
        msg += ryadok + ' Відсутній номер договору!'
    elif len ( val ) > 30:
        msg += ryadok + 'Довжина номеру договору перевищує 30 символів!'

    return msg


def val_nazva_kontragenta ( val , row_count , row ):
    global current_category
    l = len ( val )
    ryadok = 'Рядок № ' + str ( row_count ) + ':'
    msg = ''
    if val in nonstandard_categories:
        # обробляємо випадок "І.ТОВАРИ" , "ІІ.РОБОТИ" , "ІІІ.ПОСЛУГИ"
        checked_categories.append ( val )
        current_category = nonstandard_categories.index ( val ) + 1
    elif l == 0:
        msg += ryadok + ' Відсутня назва контрагента'
    elif (l > 120):
        msg += ryadok + ' Довжина назви контрагента перевищуе 120 символів.'

    return msg


def val_predmet_dogovoru ( val , row_count , row ):
    l = len ( val )
    s = str ( val )
    ryadok = 'Рядок № ' + str ( row_count ) + ':'
    msg = ''
    if l == 0:
        msg += ryadok + ' Не заповнено предмет договору'
    elif (l > 256):
        msg += ryadok + 'Довжина предмету договору перевищуе 256 символів.'

    return msg


def val_kod_cpv_rozdil ( val , row_count , row ):
    l = len ( val )
    s = str ( val )
    ryadok = 'Рядок № ' + str ( row_count ) + ':'
    msg = ''

    if l == 0:
        msg += ryadok + 'Рядок № %i : Відсутній код СPV'

    else:
        if s.find ( ' ' ) > 0 or s.find ( '"' ) > 0:
            msg += ryadok + 'Рядок № %i : Знайдені недопустимі символи-пробіл або лапки'
        elif s.count ( '-' ) > 1:
            msg += ryadok + 'Рядок № %i : Мае бути тільки один символ прочерку'
    return msg


def val_obsyag_plategiv_grn ( val , row_count , row ):
    l = len ( val )
    ryadok = 'Рядок № ' + str ( row_count ) + ':'
    msg = ''
    if l == 0:
        msg += ryadok + ' Не вказано обсяг платежів!'
    elif str ( val ).find ( ',' ) > 0:
        msg += ryadok + 'Рядок № ' + str ( row_count ) + ':'
    msg = ryadok + ' Копійки або код КВЕД в обсязи платежів повинні відокремлюватися тільки крапкою!'

    return msg


def val_vidsutni_rozdili ():
    msg = ''

    diff = list ( set ( nonstandard_categories ) - set ( checked_categories ) )
    if len ( diff ) ==0:
        msg = "Відсутніх розділів немає."
    else:
        msg = "Можливо, відсутні розділи " + str ( diff )
    return msg


# ======================================================
# встановлює зв'язок між стовпчиком та функцією для перевірки вмісту стовпчика

validators = {
    '№з/п': val_nomer_za_poryadkom ,  # тут також перевіряємо проміжні та кінцеві підсумки
    'ЄДРПОУ контрагента': val_erdpou ,
    'Договір №': val_dogovir_no ,
    'від': val_vid ,
    'Назва контрагента': val_nazva_kontragenta ,  # тут також перевіряємо категорію
    'Предмет договору': val_predmet_dogovoru ,
    'Код CPV (розділ)': val_kod_cpv_rozdil ,
    'Обсяг платежів, грн.': val_obsyag_plategiv_grn
}

# =============================================================================================

fname = get_csv_file ( )

read_from_csv ( fname )

