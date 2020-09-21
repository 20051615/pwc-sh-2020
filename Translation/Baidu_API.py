import http.client
import hashlib
import urllib
import random
import json
import time
from itertools import islice
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename

Tk().withdraw()
messagebox.showinfo("注意", "目标语言为中文，且源文件为英中混合时，会出错！")

BYTE_LIMIT = 2000
DELIMITER = "\n"

with open(askopenfilename(title="Select API_KEYS")) as keyfile:
    keyfile_content = [line.rstrip() for line in keyfile]

appid, secretKey = keyfile_content[0], keyfile_content[1]

lang_codes = {"中文": "zh", "英语": "en"}

def prompt_for_target_lang():
    langs = list(lang_codes.keys())
    for idx, lang in enumerate(langs):
        print(idx, lang, sep=": ")
    return lang_codes[langs[
        int(input("请选择目标语言（例：键入0再回车）："))]]

def fragmented_translate(toLang, source):
    # Only needed since QPS = 1; if updated API_KEY to pro/vip account, change this
    time.sleep(1)

    if len(source.encode("utf-8")) >= BYTE_LIMIT:
        print("Source is too long, please split your query.")
    else:
        httpClient = None
        myurl = '/api/trans/vip/translate'

        fromLang = 'auto'
        salt = random.randint(32768, 65536)
        sign = appid + source + str(salt) + secretKey
        sign = hashlib.md5(sign.encode()).hexdigest()
        myurl = myurl + '?appid=' + appid + '&q=' + urllib.parse.quote(source) + '&from=' + fromLang + '&to=' + toLang + '&salt=' + str(
        salt) + '&sign=' + sign

        try:
            httpClient = http.client.HTTPSConnection('api.fanyi.baidu.com')
            httpClient.request('GET', myurl)

            response = httpClient.getresponse()
            result_all = response.read().decode("utf-8")
            result = json.loads(result_all)

            return [pair['dst'] for pair in result['trans_result']]

        except Exception as e:
            print (e)
        finally:
            if httpClient:
                httpClient.close()

def isEmptyQuery(source):
    return source is None or sum(fragment != "" for fragment in source.split(DELIMITER)) == 0

# Assumption made here: no single item in source_list exceeds BYTE_LIMIT
def translate_list(toLang, source_list):
    trash = []
    trash_indices = []
    filtered_source_list = []
    for i in range(len(source_list)):
        if isEmptyQuery(source_list[i]):
            trash.append(source_list[i])
            trash_indices.append(i)
        else:
            filtered_source_list.append(source_list[i])
    source_list = filtered_source_list

    query_batches = []
    delimiter_byte_size = len(DELIMITER.encode("utf-8"))
    current_query_size = -delimiter_byte_size
    current_query_batch = []
    for source in source_list:
        size_increment = len(source.encode("utf-8")) + delimiter_byte_size
        if current_query_size + size_increment >= BYTE_LIMIT:
            query_batches.append(current_query_batch)
            current_query_size = -delimiter_byte_size
            current_query_batch = []
        current_query_batch.append(source)
        current_query_size += size_increment
    if len(current_query_batch) != 0:
        query_batches.append(current_query_batch)
    translated_list = []
    trash_to_add_index = 0
    for query_batch in query_batches:
        fragment_count_per_query = [
            sum(fragment != "" for fragment in query.split(DELIMITER))
            for query in query_batch
        ]
        translated_fragments_iterator = iter(fragmented_translate(toLang, 
                            DELIMITER.join(query_batch)))
        for fragment_count in fragment_count_per_query:
            while (trash_to_add_index < len(trash_indices)
               and
               len(translated_list) == trash_indices[trash_to_add_index]
               ):
                translated_list.append(trash[trash_to_add_index])
                trash_to_add_index += 1
            translated_list.append(DELIMITER.join(list(
                islice(translated_fragments_iterator, fragment_count))))
    while trash_to_add_index < len(trash_indices):
        translated_list.append(trash[trash_to_add_index])
        trash_to_add_index += 1
    return translated_list
