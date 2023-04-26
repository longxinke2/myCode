import requests
from cryptography.hazmat.primitives.asymmetric import rsa, padding
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives import hashes
from IPython.display import clear_output

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
    
def encode():
    public_key_bytes = b'-----BEGIN PUBLIC KEY-----\nMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzVsQAFaz1kTZXPY3Og9x\nlsspWAKXJBqv81Q4DKeL8YP1OOpihA9OYHO2um2GlVVGhlgRxbkIAEWjzelJJb8+\nSvCUm2CDwFuSYNfx7k/i5UeEK1Y7gbCjBalN2PzBuTqcvAeSfexX85Lm+lW3iQw0\nbv1k4ifKabrooq2ILJTNH3NnGjcERO8QiFa6jY5/H3eWo3POknahAX26rhpYMl1X\na/r0cSbI/c4DqzMpd6sGFQkL2DHsVCF2saY+HMPMqQHm+oT903GVAIpw0x0u6p6o\nAf1QIy8uGvZH0ee4YReLJbc80JvyHysZst3lHQu3E/0UV7nxQOPfumNfphFcwE5S\nFwIDAQAB\n-----END PUBLIC KEY-----\n'

    # 读取公钥
    public_key = serialization.load_pem_public_key(
        public_key_bytes
    )
    while True:
        with open("lib/encode.txt", "ab") as file:
            plaintext =input().encode()
            # 使用公钥加密
            ciphertext = public_key.encrypt(
                plaintext,
                padding.OAEP(
                    mgf=padding.MGF1(algorithm=hashes.SHA256()),
                    algorithm=hashes.SHA256(),
                    label=None
                )
            )
            file.write(ciphertext)
            clear_output()

def decode():
    # 读取私钥
    with open("C:/Users/龙辛柯/Desktop/code/private_key_bytes.txt", "rb") as file:
        private_key_bytes = file.read()  # 读取二进制字符串
    private_key = serialization.load_pem_private_key(
        private_key_bytes,
        password=None
    )
    with open("lib/encode.txt", "rb") as file:
        ciphertext = file.read()  # 读取二进制字符串
        text=''
        for i in range(int(len(ciphertext)/256)):
            decrypted_plaintext = private_key.decrypt(
                ciphertext[i*256:(i+1)*256],
                padding.OAEP(
                    mgf=padding.MGF1(algorithm=hashes.SHA256()),
                    algorithm=hashes.SHA256(),
                    label=None
                )
            )
            text1 = decrypted_plaintext.decode()
            if text1!='end':
                text +=text1
            else:
                text +='\n\t'
        print(text)
    with open("lib/encode.txt", "wb") as file:
        file.write(b'')
 
