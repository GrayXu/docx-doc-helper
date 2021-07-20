import re

def find_by_regex(regex, string, from_head=False):
    '''
    return a boolean to indicate that wether input string contains this regex pattern
    '''
    result = None
    if from_head:
        result = re.match(regex, string)
    else:
        result = re.search(regex, string)  # find this pattern from head
    return True if result is not None else False

def get_style_name(style):
    return str(style).split("\'")[1]

def replace_para_text(para, source, target):
    '''
    替换para的字体，但不修改字体（取最后一个run）
    '''
    # method below will break style!
    # para.text = str(para.text).replace(source, target)
    para_text = str(para.text).replace(source, target)
    for run in para.runs:
        run.text = ''
    para.runs[-1] = para_text
    
    
def get_sec_num(ori):
    '''
    get section number:   '23.1 haha' -> '23.1'
    '''
    ori = ori.strip()
    flag = False  # a flag status for pattern detect
    
    if len(ori) != 0 and ori[0].isdigit():  # first char is a digit
        for char_index in range(len(ori)-1):
            char = ori[1+char_index]  # jump the first one
            if char == ' ' and flag:  # find space end
                return ori[:char_index+1]
            elif char == '.':
                flag = True
            elif char.isdigit():  # keep move
                continue
            else:  # a normal char
                if flag:
                    return ori[:char_index+1]
                else: 
                    break
    return None

def is_sec_pattern(ori):
    '''
    A string is in section title pattern
    '''
    ori = ori.strip()
    if len(ori) != 0 and ori[0].isdigit():  # first char is a digit
        for char in ori[1:]:
            if char == '.':
                return True
            elif char.isdigit():
                continue
            else:
                break
    return False

