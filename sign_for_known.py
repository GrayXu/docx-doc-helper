'''
2021/7/6
author: grayxu

批量采纳已知pattern的意见（“必选的”等）
'''

import csv

known_pattern = set(['必选的','可选的','有条件的'])

found = []
data = []
with open('desktop/7.6.csv') as file:
    reader = csv.reader(file,delimiter=',')
    count = 0
    for row in reader:
        count += 1
        if row[3] in known_pattern:
            found.append(count)
            row[5] = '采纳'
        data.append(row)
        
with open('desktop/7.6_n.csv', 'w') as file:
    writer = csv.writer(file)
    writer.writerows(data)
    
print('finished')