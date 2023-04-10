import requests
def translate():
    s = input(":")
    dat = {
        "kw":s
    }
    resp = requests.post("https://fanyi.baidu.com/sug", data=dat)
    print(resp.json()['data'][0]['v']) #将返回的内容直接返回为json

def is_English():
    string = input(":")
    for ch in string:
        if ord(ch) >= 128:
            return ch
    return True

def progress_bar(now,end):
    percent = now*100//end
    line = '-'*(percent//2)+' '*(50-percent//2)
    print(f'\r{line}{percent}%',end="")
