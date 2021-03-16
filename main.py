#!/home/sasha/dev/Python/wialonAPI/venv python3.8


from wialon import Wialon, WialonError
import xlsxwriter as xls
import math as m
import datetime
import logistics_token


def sec_to_time(sec):
    time_sec = int(m.fmod(sec, 60))
    time_min = int(m.fmod(sec//60, 60))
    time_hour = int(m.fmod((sec//60)//60, 24))
    return '{HH:02.0f}:{mm:02.0f}:{ss:02.0f}'.format(HH=time_hour, mm=time_min, ss=time_sec)


if __name__ == '__main__':
    wialon_api = Wialon()
    for token in logistics_token.token:
        result = wialon_api.token_login(token=token)
        wialon_api.sid = result['eid']
        spec = {
            "itemsType": "avl_resource",
            "propType": "propitemname",
            "propName": "orders",
            "propValueMask": "*",
            "sortType": "orders"
        }
        params = {
            "spec": spec,
            "force": 1,
            "flags": 524288,
            "from": 0,
            "to": 0
        }
        orders = wialon_api.call('core_search_items', params)
        orders = orders['items']
        workbook = xls.Workbook('export orders/{:%Y-%m-%d}-{}.xlsx'.format(datetime.datetime.now(), result['au']))
        worksheet = workbook.add_worksheet()
        str_xlsx = 0
        worksheet.write(str_xlsx, 0, '№')
        worksheet.write(str_xlsx, 1, 'Id')
        worksheet.write(str_xlsx, 2, 'Заявка')
        worksheet.write(str_xlsx, 3, 'Клиент')
        worksheet.write(str_xlsx, 4, 'Адрес')
        worksheet.write(str_xlsx, 5, 'Стоимость')
        worksheet.write(str_xlsx, 6, 'Сервистное время')
        worksheet.write(str_xlsx, 7, 'Теги')
        worksheet.write(str_xlsx, 8, 'Время от')
        worksheet.write(str_xlsx, 9, 'Время до')
        worksheet.write(str_xlsx, 10, 'Радиус')
        worksheet.write(str_xlsx, 11, 'Долгота')
        worksheet.write(str_xlsx, 12, 'Широта')
        custom_filds = dict()
        number_cf = 12
        for el in orders:
            orders = el['orders']
            for num in orders.values():
                if num['f'] == 32:
                    str_xlsx += 1
                    worksheet.write(str_xlsx, 0, str_xlsx)
                    worksheet.write(str_xlsx, 1, num['id'])
                    worksheet.write(str_xlsx, 2, num['n'])
                    worksheet.write(str_xlsx, 3, num['p']['n'])
                    worksheet.write(str_xlsx, 4, num['p']['a'])
                    worksheet.write(str_xlsx, 5, num['p']['c'])
                    st = int(num['p']['ut']//60)
                    worksheet.write(str_xlsx, 6, st)
                    worksheet.write(str_xlsx, 7, '|'.join(num['p']['tags']))
                    worksheet.write(str_xlsx, 8, sec_to_time(num['tf']))
                    worksheet.write(str_xlsx, 9, sec_to_time(num['tt']))
                    worksheet.write(str_xlsx, 10, num['r'])
                    worksheet.write(str_xlsx, 11, num['y'])
                    worksheet.write(str_xlsx, 12, num['x'])

                    if num['cf'] is not None:
                        cf = num['cf']
                        for key, values in cf.items():
                            if custom_filds.get(key) is None:
                                number_cf += 1
                                custom_filds.update({key: number_cf})
                                worksheet.write(0, custom_filds.get(key), key)
                            worksheet.write(str_xlsx, custom_filds.get(key), values)
        workbook.close()
        wialon_api.core_logout()
    print('Ok')

