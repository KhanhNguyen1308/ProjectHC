def DateEdit(date):
    if len(date) == 3:
        if len(date[1]) < 2:
            date[1] = "0"+str(date[1])
        if len(date[0]) < 2:
            date[0] = "0"+str(date[0])
        if len(date[2]) > 2:
            date[2] = date[2][-2:]
        date_text = str(date[0])+"/"+str(date[1])+"/"+str(date[2])
    else:
        date_text='               '
    return date_text


