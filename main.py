# -*- coding:utf-8 -*-
import requests

def get_topic_post():
    url_diary = 'https://web.okjike.com/topic/5aeaa84029e4000011ac3768'
    url_memo = 'https://web.okjike.com/topic/5628fac0daf87d13002c8964'
    def show(url):
        response = requests.get(url)
        row_text = response.text
        get_js = row_text.split('<script')[-1]  #  提取JS部分，因为body里面的东西是用js渲染的
        get_content_all = get_js.split('content":')[2:]  # 提取其中的content， content后面还有很多其他参数，其中第一个是 shareCount ,再从这里断开就好了
        content_list = []
        for item in get_content_all:
            content_list.append(item.split(',"shareCount"')[0])
        return content_list

    for i in show(url_memo):
        print(i)
    return 'down'

get_topic_post()