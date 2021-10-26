import telebot, os, os.path

bot = telebot.TeleBot("2036737833:AAGUnUw0eSz6G3FQkBf5CU9HQQBWJRzJ1WU")

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.send_message(message.chat.id, "Привет ленивая задница. \nЗагрузи файл.. (.xls)")


@bot.message_handler(content_types=['document'])
def handle_docs(message):
    file_name = message.document.file_name
    file_id_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_id_info.file_path)
    with open(file_name, 'wb') as new_file:
        new_file.write(downloaded_file)
    bot.send_message(message.chat.id, "Файл успешно получен. Ожидаете модерации...")

    extension = os.path.splitext(file_name)[1]
    if(extension != '.xls'):
        bot.send_message(message.chat.id, "Переданный файл поврежден или не являеться отчётом 1С. Повторите попытку.")
        os.remove(file_name)
    else:
        import xlrd, xlwt

        def toFixed(numObj, digits=2):
            return f"{numObj:.{digits}f}"




        fields = ['№', 'Залоговый билет', 'Клиент %', '%', 'Телефон', 'Ссуда', '% (нач.)', '% (опл.)', 'Кол. дней', 'К оплате', 'Suma lun.', 'Datoria', 'Datoria(L)']

        align = xlwt.Alignment()
        align.horz = 0x02
        align.vert = 0x01

        borders = xlwt.Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1

        fieldPattern = xlwt.Pattern()
        fieldPattern.pattern = xlwt.Pattern.SOLID_PATTERN
        fieldPattern.pattern_fore_colour = xlwt.Style.colour_map['gray25']

        fontTitle = xlwt.Font()
        fontTitle.name = 'Calibri'
        fontTitle.height = 480
        fontTitle.bold = True

        fontField = xlwt.Font()
        fontField.name = 'Calibri'
        fontField.height = 200
        fontField.bold = True

        fontCell = xlwt.Font()
        fontCell.name = 'Calibri'
        fontCell.height = 140

        titleStyle = xlwt.XFStyle()
        titleStyle.font = fontTitle
        titleStyle.alignment = align


        fieldStyle = xlwt.XFStyle()
        fieldStyle.font = fontField
        fieldStyle.alignment = align
        fieldStyle.pattern = fieldPattern
        fieldStyle.borders = borders

        cellStyle = xlwt.XFStyle()
        cellStyle.font = fontCell
        cellStyle.alignment = align
        cellStyle.borders = borders




        xlsI = xlrd.open_workbook(file_name, encoding_override="cp1251")
        iSheet = xlsI.sheet_by_index(0)

        tableLen = 0
        while "Итого" not in str(iSheet.cell_value(tableLen, 1)):
            tableLen += 1
        table = []

        for i in range(5, tableLen):
            row = ['','','','','','','','','','','','','']
            row[0] = str(iSheet.cell_value(i, 1))
            row[1] = str(iSheet.cell_value(i, 3))
            row[2] = str(iSheet.cell_value(i, 12))
            row[3] = str(iSheet.cell_value(i, 18))
            row[4] = str(iSheet.cell_value(i, 19))
            row[5] = str(iSheet.cell_value(i, 25))
            row[6] = str(iSheet.cell_value(i, 30))
            row[7] = str(iSheet.cell_value(i, 34))
            row[8] = str(iSheet.cell_value(i, 38))
            row[9] = str(iSheet.cell_value(i, 42))
            row[10] = toFixed(float(row[5])*float(row[3])/100)
            row[11] = toFixed(float(row[9]) - float(row[5]))
            row[12] = toFixed(float(row[11]) / float(row[10]))
            if float(row[12]) > 1:
                table.append(row)

        for i in table:
            i[12] = float(i[12])

        table = sorted(table, key=lambda tup: tup[12], reverse=True)


        xlsO = xlwt.Workbook(encoding="cp1251")
        oSheet = xlsO.add_sheet('Отчёт', cell_overwrite_ok = True)

        oSheet.col(0).width = 5*256
        oSheet.col(1).width = 20*256
        oSheet.col(2).width = 22*256
        oSheet.col(3).width = 3*256
        oSheet.col(4).width = 20*256
        oSheet.col(5).width = 9*256
        oSheet.col(6).width = 9*256
        oSheet.col(7).width = 9*256
        oSheet.col(8).width = 9*256
        oSheet.col(9).width = 9*256
        oSheet.col(10).width = 9*256
        oSheet.col(11).width = 9*256
        oSheet.col(12).width = 9*256


        tall_style = xlwt.easyxf('font:height 50;')
        for i in range(len(table)+1):
            oSheet.row(i+4).set_style(tall_style)

        #установка Title
        title = str(iSheet.cell_value(1, 1)) + str(iSheet.cell_value(2, 1))
        oSheet.write_merge(0, 3, 0, 12, title, style=titleStyle)

        #установка Fields
        for field in fields:
            oSheet.write(4,fields.index(field), field, style=fieldStyle)

        #установка Cell
        for i in range(len(table)):
            for row in table:
                oSheet.write(i+5, 0, toFixed(float(table[i][0]),0), cellStyle)
                oSheet.write(i+5, 1, table[i][1], cellStyle)
                oSheet.write(i+5, 2, table[i][2], cellStyle)
                oSheet.write(i+5, 3, table[i][3], cellStyle)
                oSheet.write(i+5, 4, table[i][4], cellStyle)
                oSheet.write(i+5, 5, table[i][5], cellStyle)
                oSheet.write(i+5, 6, table[i][6], cellStyle)
                oSheet.write(i+5, 7, table[i][7], cellStyle)
                oSheet.write(i+5, 8, table[i][8], cellStyle)
                oSheet.write(i+5, 9, table[i][9], cellStyle)
                oSheet.write(i+5, 10, table[i][10], cellStyle)
                oSheet.write(i+5, 11, table[i][11], cellStyle)
                oSheet.write(i+5, 12, table[i][12], cellStyle)

        os.remove(file_name)
        xlsO.save('Отчёт - ' + file_name)
        bot.send_message(message.chat.id, "Отчёт готов. Жду тебя на следующей недели.")
        f = open('Отчёт - ' + file_name, 'rb')
        bot.send_document(message.chat.id, f)
        f.close()
        os.remove('Отчёт - ' + file_name)



bot.polling()