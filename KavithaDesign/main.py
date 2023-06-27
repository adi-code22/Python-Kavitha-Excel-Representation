import xlsxwriter

print('testing. . . .')

wb = xlsxwriter.Workbook('hello.xlsx')

# colors *********************************************
greenColor1 = wb.add_format()
greenColor2 = wb.add_format()
greenColor3 = wb.add_format()
greenColor4 = wb.add_format()

greenColor1.set_bg_color('#4CBB17')
greenColor2.set_bg_color('#0BDA51')
greenColor3.set_bg_color('#90EE90')
greenColor4.set_bg_color('#AFE1AF')
# ****************************************************

# positions ******************************************
word1Pos = ['A4', 'B3', 'C2', 'D1', 'E1', 'F2', 'G3', 'H4']
word2Pos = ['B5', 'B4', 'C3', 'D2', 'E2', 'F3', 'G4', 'G5']
word3Pos = ['C6', 'C5', 'C4', 'D3', 'E3', 'F4', 'F5', 'F6']
word4Pos = ['D7', 'D6', 'D5', 'D4', 'E4', 'E5', 'E6', 'E7']

word1index = word2index = word3index = word4index = 0
# *************************************************

ws = wb.add_worksheet()

for i in range(0, 7):
    ws.set_row(i, 60)

for i in range(65, 73):
    ws.set_column(chr(i) + ':' + chr(i), 10)

# ws.write('A1', 'Hello world', greenColor1)
malGrp = ['ഀ', 'ഁ', 'ം', 'ഃ', 'ഄ', '഼', '഻', 'ീ', 'ു', 'ൂ', 'ൃ', 'ൄ', 'െ', 'േ', 'ൈ', 'ൊ', 'ോ', 'ൌ', '്', 'ൎ', 'ൗ', 'ാ', 'ി']

word1 = input("Enter 1st word: ")
prev = ''
for i in range(0, len(word1)):
    if not malGrp.__contains__(word1[i]):
        prev = word1[i]
        if i == len(word1) - 1 and word1index < 8:
            ws.write(word1Pos[word1index], word1[i], greenColor1)
            word1index += 1
        elif i < len(word1) and word1index < 8 and not malGrp.__contains__(word1[i+1]):
            ws.write(word1Pos[word1index], word1[i], greenColor1)
            word1index += 1
    elif malGrp.__contains__(word1[i]):
        val = prev + word1[i]
        if word1index < 8:
            ws.write(word1Pos[word1index], val, greenColor1)
            word1index += 1

word2 = input("Enter 2nd word: ")
prev = ''
for i in range(0, len(word2)):
    if not malGrp.__contains__(word2[i]):
        prev = word2[i]
        if i == len(word2) - 1 and word2index < 8:
            ws.write(word2Pos[word2index], word2[i], greenColor2)
            word2index += 1
        elif i < len(word2) and word2index < 8 and not malGrp.__contains__(word2[i+1]):
            ws.write(word2Pos[word2index], word2[i], greenColor2)
            word2index += 1
    elif malGrp.__contains__(word2[i]):
        val = prev + word2[i]
        if word2index < 8:
            ws.write(word2Pos[word2index], val, greenColor2)
            word2index += 1

word3 = input("Enter 3rd word: ")
prev = ''
for i in range(0, len(word3)):
    if not malGrp.__contains__(word3[i]):
        prev = word3[i]
        if i == len(word3) - 1 and word3index < 8:
            ws.write(word3Pos[word3index], word3[i], greenColor3)
            word3index += 1
        elif i < len(word3) and word3index < 8 and not malGrp.__contains__(word3[i+1]):
            ws.write(word3Pos[word3index], word3[i], greenColor3)
            word3index += 1
    elif malGrp.__contains__(word3[i]):
        val = prev + word3[i]
        if word3index < 8:
            ws.write(word3Pos[word3index], val, greenColor3)
            word3index += 1

word4 = input("Enter 4th word: ")
prev = ''
for i in range(0, len(word4)):
    if not malGrp.__contains__(word4[i]):
        prev = word4[i]
        if i == len(word4) - 1 and word4index < 8:
            ws.write(word4Pos[word4index], word4[i], greenColor4)
            word4index += 1
        elif i < len(word4) and word4index < 8 and not malGrp.__contains__(word4[i+1]):
            ws.write(word4Pos[word4index], word4[i], greenColor4)
            word4index += 1
    elif malGrp.__contains__(word4[i]):
        val = prev + word4[i]
        if word4index < 8:
            ws.write(word4Pos[word4index], val, greenColor4)
            word4index += 1

print(word1 + " " + word2 + " " + word3 + " " + word4)

# hiding unwanted rows and columns
ws.set_default_row(hide_unused_rows=True)
ws.set_column("I:XFD", None, None, {"hidden": True})

wb.close()