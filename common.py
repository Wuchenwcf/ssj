
import codecs
import chardet

def detect_encoding(file_name):
    with open(file_name, 'rb') as file:
        rawdata = file.read()
    result = chardet.detect(rawdata)
    return result['encoding']




def remove_lines_range(file_name, start, end):
    with open(file_name, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    with open(file_name, 'w', encoding='utf-8') as file:
        #print(lines[:max(start-1,0)])
        #print(lines[end:])
        file.writelines(lines[:max(start,1)-1] + lines[end:])

if __name__ == '__main__':
    # 使用示例
    file_name = 'test.csv'
    encoding = detect_encoding(file_name)
    print(f"The encoding of {file_name} is: {encoding}")

    #convert_utf8_bom_to_utf8(file_name)
    start_line = 0  # 起始行（包含）
    end_line = 21    # 结束行（包含）
    remove_lines_range(file_name, start_line, end_line)