import json
import base64
import os

def load(path,file_name):
    path = path+ "\\"+file_name
    if os.path.isfile(path) is True:
        with open(path, 'r') as json_file:
            data = json.load(json_file)
            return data
    else:
        return None

def store(data,path,file_name):
    if os.path.isdir(path) is False:
        os.mkdir(path)
    path = path + "\\"+file_name
    with open(path, 'w') as json_file:
        json_file.write(json.dumps(data))

def base64_encode(data):
    keys=list(data.keys())
    value_encode_list = []
    key_encode_list =[]
    for key in keys:
        value_encode = base64.b64encode(data[key].encode("utf-8"))
        value_encode_list.append(value_encode)
        key_encode_list.append(base64.b64encode(key.encode("utf-8")))
    data = dict(zip(key_encode_list, value_encode_list))
    return data

def base64_decode(data):
    keys_encode_list = list(data.keys())
    values_encode_list =list(data.values())
    keys = []
    values =[]
    for key in keys_encode_list:
        key_encode = base64.b64decode(key)
        keys.append(key_encode.decode("utf-8"))
    for value in values_encode_list:
        values.append(base64.b64decode(value).decode("utf-8"))
    data = dict(zip(keys, values))
    return(data)

def load_with_decode(path,file_name):
    path = path + "\\" + file_name
    if os.path.isfile(path) is True:
        with open(path, 'r') as file:
            data = base64_decode(eval(file.read()))
        return data
    else:
        return None

def store_with_encode(data,path,file_name):
    if os.path.isdir(path) is False:
        os.mkdir(path)
    path = path + "\\"+file_name
    data = base64_encode(data)
    with open(path, 'w') as file:
        file.write(str(data))
    return data

if __name__ == "__main__":
    # data = load(path="D:",file_name="data.jason")
    # print(data)
    data = {
        "user_name":"joe"
    }
    data = store_with_encode(data=data,path = "D:",file_name="data.jason")
    print("写入成功")
    print(data)
    data = load_with_decode(path = "D:",file_name="data.jason")
    print(data)

    # print(data,"\n",str(data))

    # # print(data)
    # print(data)

