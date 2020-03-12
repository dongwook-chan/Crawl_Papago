import base64
import re
import requests

import json
import openpyxl

class Translator:
    """
    Title: Pypago
    Author: Beomi@GitHub
    Date: 8 Jul 2019
    Code version: 0.1.1.1
    Availability: https://github.com/Beomi/pypapago/blob/0.1.1.1/pypapago/translator.py
    """

    def __init__(self, regex_pattern=None, headers=None):
        self.regex_pattern = re.compile(regex_pattern or '[가-힣]+')
        self.headers = headers or {
            'device-type': 'pc',
            'origin': 'https://papago.naver.com',
            'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'ko',
            'authority': 'papago.naver.com',
            'pragma': 'no-cache',
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_4) AppleWebKit/537.36 (KHTML, like Gecko)\
                           Chrome/75.0.3770.100 Safari/537.36',
            'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'accept': 'application/json',
            'cache-control': 'no-cache',
            'x-apigw-partnerid': 'papago',
            'referer': 'https://papago.naver.com/',
            'dnt': '1',
        }
        self.SECRET_KEY = 'rlWxMKMcL2IWMPV6ImUwMWMwZWFkLWMyNDUtNDg2YS05ZTdiLWExZTZmNzc2OTc0MyIsImRpY3QiOnRydWUsImRpY3REaXNwbGF5Ijoz'
        self.QUERY_KEY = '0,"honorific":false,"instant":false,"source":"{source}","target":"{target}","text":"{query}"}}'

    @staticmethod
    def string_to_base64(s):
        """
        Generate Base64 Encoded string
        :param s: Origin Text (UTF-8)
        :return: B64 encoded text (B64, still UTF-8 string)
        """
        return base64.b64encode(s.encode('utf-8')).decode('utf-8')

    def translate(self, query, source='ko', target='en', verbose=True):
        """
        Main Translate function
        :param query: Original Text to translate
        :param source: Source(Original) text language [en, ko]
        :param target: Target text language [en, ko]
        :param verbose: Return verbose json data. Default: False
        :return: Translated text
        """
        data = {
            'data': self.SECRET_KEY + self.string_to_base64(
                self.QUERY_KEY.format(source=source, target=target, query=query)
            )
        }
        response = requests.post('https://papago.naver.com/apis/n2mt/translate', headers=self.headers, data=data)
        if not verbose:
            return response.json()['translatedText']
        return response.json()

translator = Translator();
wb = openpyxl.load_workbook('input.xlsx')
ws = wb['Sheet1']
word_dict = dict()

with open('index.txt', 'r', encoding='utf-8') as inFile:
    indexStr = inFile.read()
    index = int(indexStr)

for r in ws['C'+ indexStr :'C5966']:
    word = translator.translate(r[0].value)
    print(str(index) + ' ' + r[0].value)
    with open(str(index) + '.json', 'w', encoding='utf-8') as outFile1:
        json.dump(word, outFile1, indent="\t")
    index = index + 1
    with open('index.txt', 'w', encoding='utf-8') as outFile2:
        outFile2.write(str(index))
