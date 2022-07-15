from time import sleep
import requests
import xlwt

from time import time
import xlwt
from random import randint
from hashlib import md5

def tprint(message, tip='tip'):  # 格式化提示信息
    if tip == '':
        print(message)
    else:
        START_LINE = '\n=-- ' + tip + ' --:'
        print(START_LINE)
        print(message)
        END_LINE = ''
        for _ in START_LINE:
            END_LINE += '-'
        END_LINE = END_LINE[:-2] + '='
        print(END_LINE)


class Excel():
    def __init__(self, exl_name = 'new_', sheet_name=['Sheet1'], start_line = 0 , start_col = 0, type = '.xls', json_data={}, exl_path='./'):
        self.type = type
        self.exl_name = exl_name
        self.start_line = start_line
        self.start_col = start_col
        self.workbook = xlwt.Workbook(encoding= 'utf-8')
        self.worksheet = self.workbook.add_sheet(sheet_name[0], cell_overwrite_ok=True)
        self.BASE = {
            'x' : self.start_col,
            'y' : self.start_line
                }
        self.row_mark = 0  # 相对BASE line 的距离
        self.col_mark = 0
        self.json_data = json_data
        self.exl_path = exl_path

    @staticmethod
    def font_style( name='宋体',
                    bold=False,
                    italic=False,
                    underline=False,
                    colour_index='black',
                    height=200
        ):
        font = xlwt.Font()
        font.name = name
        font.bold = bold
        font.underline = underline
        font.italic = italic
        font.colour_index = xlwt.Style.colour_map[colour_index]
        font.height = height  # 200为12号字体 400为20号字体
        return font

    # 设置并返回,边框对象
    @staticmethod
    def border_style(left=1, right=1, top=1, bottom=1):
        borders = xlwt.Borders()
        borders.left = left
        borders.right = right
        borders.top = top
        borders.bottom = bottom
        return borders

    # 设置并返回,对其样式对象
    @staticmethod
    def align_style(horz=True, vert=True):
        align = xlwt.Alignment()
        if horz=='Right':
            align.horz = xlwt.Alignment.HORZ_RIGHT  # 水平方向
        elif horz==True:
            align.horz = xlwt.Alignment.HORZ_CENTER
        else:
            align.horz = xlwt.Alignment.HORZ_LEFT
        if vert:
            align.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
        align.wrap = 1
        return align

    # 设置并返回,背景颜色样式对象
    @staticmethod
    def set_BG_color(color='white'):  # 设置单元格背景颜色
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = xlwt.Style.colour_map[color]  # 设置单元格背景色为绿色
        return pattern

    def line_mark_add_1(self):
        self.row_mark += 1

    def set_width(self, colx, pace):  # 改变指定列的间距
        # 40意味着40磅
        if colx >= 0:
            self.worksheet.col(self.BASE['y'] + colx).width_mismatch = True
            self.worksheet.col(self.BASE['y'] + colx).width = pace

    # 改变当前实例指定行的间距
    # 若不指定行, 则改变当前游标所在行的间距
    def set_height(self, pace, rowx=0):
        if rowx > 0:
            self.worksheet.row(self.BASE['x'] + rowx).height_mismatch = True  # 默认行高是和文字的高度进行匹配的,即 height_mismatch
            self.worksheet.row(self.BASE['x'] + rowx).height = pace
        elif rowx == 0 :
            rowx = self.row_mark
            self.worksheet.row(self.BASE['x'] + rowx).height_mismatch = True  # 默认行高是和文字的高度进行匹配的,即 height_mismatch
            self.worksheet.row(self.BASE['x'] + rowx).height = pace
        else:
            return 'INVALID ROWX.'

    # 在当前实例中进行表格的合并操作
    def cells_merge(self, cell_1, cell_2, content='', decimal=0, align_horz=True,  if_bold = False, fontsize=200, font_color_index='black' ,bg_color='',border=[]):  # cell_1 => (x1,y1) , cell_2 => (x2,y2)
        style = xlwt.XFStyle()
        style.alignment = self.align_style(horz=align_horz)
        style.font = self.font_style(bold=if_bold, colour_index=font_color_index, height=fontsize)

        if bg_color != '':
            style.pattern = self.set_BG_color(bg_color)
        if len(border) > 0:
            borders= xlwt.Borders()
            borders.left= border[0]
            borders.right= border[1]
            borders.top= border[2]
            borders.bottom= border[3]
            style.borders = borders

        if decimal > 0 and decimal < 10:
            num_format_str = '0.'
            for _ in range(decimal):
                num_format_str += '0'
            style.num_format_str = num_format_str
        self.worksheet.write_merge( self.BASE['x'] + cell_1[0],
                                    self.BASE['x'] + cell_2[0],
                                    self.BASE['y'] + cell_1[1],
                                    self.BASE['y'] + cell_2[1],
                                    content, style)

    # 改变某一格的内容
    def set_value(self, bg_color='', align_horz=True,x='', y='', value='', if_bold = False, font_color_index='black', fontsize=200, border=[], decimal=0):  # 当不指定x时, 将在当前line mark所指定行 添加数据
        if x == '':
            x = self.row_mark
        if bg_color!= '' or len(border) > 0 or (decimal > 0 and decimal < 10):
            style = xlwt.XFStyle()
            style.alignment = self.align_style(horz=align_horz)

            style.font = self.font_style(bold=if_bold, colour_index=font_color_index, height=fontsize)
            style.font = self.font_style(bold=if_bold, colour_index=font_color_index, height=fontsize)
            if bg_color != '':
                style.pattern = self.set_BG_color(bg_color)
            if len(border) > 0:
                borders= xlwt.Borders()
                borders.left= border[0]
                borders.right= border[1]
                borders.top= border[2]
                borders.bottom= border[3]
                style.borders = borders

            if decimal > 0 and decimal < 10:
                num_format_str = '0.'
                for i in range(decimal):
                    num_format_str += '0'
                style.num_format_str = num_format_str

            self.worksheet.write(self.BASE['x'] + x, self.BASE['y'] + y, value, style)
        else:
            self.worksheet.write(self.BASE['x'] + x, self.BASE['y'] + y, value)

    # 在当前实例中设置表头样式
    def build_head( self,
                    cell_1,
                    cell_2,
                    text,
                    borderstyle='',
                    fontstyle='',
                    BGcolor='',
                    alignstyle='',
                    width=0,
                    height=0):  # 生成一个合并的单元格,作为表头
        # 起、止的行、列, 需要合并的单元格;  内容;  字号 ; 字体;  对其方式
        '''
        params:
            content
            size
            style
            color
            align
        '''
        style = xlwt.XFStyle()

        # 边框
        if not borderstyle:
            style.borders = self.border_style()
        else:
            style.borders = borderstyle

        # 字体
        if not fontstyle:
            style.font = self.font_style()
        else:
            style.font = fontstyle

        # 单元格居中
        if not alignstyle:
            style.alignment = self.align_style()
        else:
            style.alignment = alignstyle

        # 背景色
        if not BGcolor:
            style.pattern = self.set_BG_color()
        else:
            style.pattern = self.set_BG_color(BGcolor)

        # 表头样式

        self.worksheet.write_merge( cell_1[0],
                                    cell_2[0],
                                    cell_1[1],
                                    cell_2[1],
                                    text, style)

        # 设置表的基准参照
        self.BASE['x'] = cell_1[0]
        self.BASE['y'] = cell_1[1]
        for i in range(cell_1[0]):
            self.line_mark_add_1()
        self.col_mark = cell_1[1]

        # # 长宽 size
        if width:
            self.set_width(x = cell_1[1], pace=width)
        if height:
            self.worksheet.row(cell_1[0]).height_mismatch = True  # 默认行高是和文字的高度进行匹配的,即 height_mismatch
            self.worksheet.row(cell_1[0]).height = height

    def build_foot(self, param_list={}):  # 生成一个合并的单元格作为,表尾。
        pass

    def insert_bmp(self, url, x, y, x1=0, y1=0, scale_x=1, scale_y=1):  # 指定位置插入图片
        '''
        x 表示行, y 表示列x1 y1 表示相对原来位置向下向右偏移的像素scale_x scale_y表示相对原图宽高的比例, 图片可放大缩小
        '''
        self.worksheet.insert_bitmap(url, x, y, x1, y1, scale_x, scale_y)

    def build_col_head(self):
        pass

    def add_block(self):
        pass

    def build_foot(self):
        pass

    def generate_excel(self):
        # 添加标题行
        self.set_value(value="日期",x=self.row_mark,y=0,border=[1,1,1,2])  # 企业名称
        self.set_width(0,10000)
        self.set_value(value="内容",x=self.row_mark,y=1,border=[1,1,1,2])  # 法人名称
        self.set_width(1,60262)
        self.line_mark_add_1()
        for key in self.json_data:
            i = self.json_data[key]
            self.set_value(value=i['date'],x=self.row_mark,y=0)  # 企业名称
            self.set_value(value=i['content'],x=self.row_mark,y=1)  # 法人名称
            self.set_height(300)
            self.line_mark_add_1()
        self.save()

    # 将当前实例中的 Workbook 保存到具体的路径
    def save(self, exl_path = './'):  # 在指定位置保存excel文件
        try:
            info  = 'The file {name} saved at {path}.'.format(name = self.exl_name, path = exl_path)
            self.workbook.save(exl_path + self.exl_name + self.type)
            tprint(info, 'info')
        except Exception as e:
            tprint(e, 'err')
        pass



def build_excel(JSON_DATA, exl_path='./',excel_name_three=''):
    try:
        RAWNAME = excel_name_three
        test_excel_3 = Excel(exl_name=RAWNAME, json_data=JSON_DATA,exl_path=exl_path)
        excelname_3 = test_excel_3.generate_excel()
        return excelname_3

    except Exception as e:
        raise(e)
        return 'Build Faild'




class jike():
    def __init__(self) -> None:
        self.headers = {
            "accept": "*/*",
            "accept-language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
            "accept-encoding": "gzip, deflate, br",
            "Content-Length": "6854",
            "content-type":"application/json",
            "cookie": "_ga=GA1.2.776104600.1650769729; _gid=GA1.2.2087473452.1657684172; fetchRankedUpdate=1657850949355; x-jike-access-token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJkYXRhIjoibkhrRVFhbFUzZ3NVWmNpQkJidTNPQ2hpbk1LZHkwNUdEM0Y1VHFRNjM2TzF2ZHRydzVqc1FYS0dzaG4wR0tBRlNCeEx1dXNDYUNlN0pOMERBRHZxK3FZajNhc3ZJNmZSY2dVc08rS0N5Z3ZoSUJtTUdcL2JnWG5BRHNseGJtVTRWYzZNaTg5NllTUmNWR09sTm43MzNIOGp1WVBGRHJEbHRVdzl1UG9IMkc2WXZONStLQ1U0Q1hhcEhcL3UyMXNndFNvZHN0Tmh3YXF0dEd1VjQrUmxlYXlsOFRnYWlReHJCcjRBQjZNWlNtWWk5eDhmZE9VMXVXaHVKSVcwVDltZGNXQ3lRa3N2SGpiYzJTQ2JHbFdiTmgzK2J4SUpMUURJWnFhWG9pSnd2VW9STWNZOWlSQU9VK2RiXC9wWlN5SkxWdXVTZWlTeDVTdE9uXC83dFlXT2JRelVQaWdLNHBRY0ZUT2VlUlBhV3QyTnNjaz0iLCJ2IjozLCJpdiI6IlZIZXdZUktFTHFKWTVKMFI3XC9cL29VZz09IiwiaWF0IjoxNjU3ODY5Njg5LjkyNH0.uYZBjEFz1fa8NrYisOqGtwckdLNDXu4sSCpxvIN8Vlo; x-jike-refresh-token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJkYXRhIjoiN1wvWTByZ0Y3cTdYRGNzbGFJVGs3MHdSOG1xRTFHUWJPZHZ1OEcrQVcxQUY2ZlBDazlldTdxek9oOU95Q1NOcWZsZ0dcL2piTDNtRzdpRFwvVFFHelNIXC83bW84ZjZcL2Z6azRNWXFZaXNNNWYrcGpYbnhIVFdUV29LUTBhaTZYQWx2cGtHMGx4TllWSk96bmc0bDJkSDV2TWlRRHRwYkNZTlZQYURwYk45d2pJQ3M9IiwidiI6MywiaXYiOiIxR0FKbFFaMHl6aG9nVDRkVUZqdmFnPT0iLCJpYXQiOjE2NTc4Njk2ODkuOTI0fQ.qKs6Zr-aLDHAiZK5nVy3tFwqO0XGs34yIJ8dh2fBNcc",
            "origin": "https://web.okjike.com",
            "sec-ch-ua": '".Not/A)Brand";v="99", "Google Chrome";v="103", "Chromium";v="103"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "Windows",
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "same-site",
            "dnt":"1",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"
        }

    def show_all_post(self,__username):
        json = {
        "operationName":"UserFeeds",
        "variables":{
            "username":__username,
        },
        "query":"query UserFeeds($username: String!, $loadMoreKey: JSON) {\n  userProfile(username: $username) {\n    username\n    feeds(loadMoreKey: $loadMoreKey) {\n      ...BasicFeedItem\n      __typename\n    }\n    __typename\n  }\n}\n\nfragment BasicFeedItem on FeedsConnection {\n  pageInfo {\n    loadMoreKey\n    hasNextPage\n    __typename\n  }\n  nodes {\n    ... on ReadSplitBar {\n      id\n      type\n      text\n      __typename\n    }\n    ... on MessageEssential {\n      ...FeedMessageFragment\n      __typename\n    }\n    ... on UserAction {\n      id\n      type\n      action\n      actionTime\n      ... on UserFollowAction {\n        users {\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          __typename\n        }\n        allTargetUsers {\n          ...TinyUserFragment\n          following\n          statsCount {\n            followedCount\n            __typename\n          }\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          __typename\n        }\n        __typename\n      }\n      ... on UserRespectAction {\n        users {\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          __typename\n        }\n        targetUsers {\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          ...TinyUserFragment\n          __typename\n        }\n        content\n        __typename\n      }\n      __typename\n    }\n    __typename\n  }\n  __typename\n}\n\nfragment FeedMessageFragment on MessageEssential {\n  ...EssentialFragment\n  ... on OriginalPost {\n    ...LikeableFragment\n    ...CommentableFragment\n    ...RootMessageFragment\n    ...UserPostFragment\n    ...MessageInfoFragment\n    pinned {\n      personalUpdate\n      __typename\n    }\n    __typename\n  }\n  ... on Repost {\n    ...LikeableFragment\n    ...CommentableFragment\n    ...UserPostFragment\n    ...RepostFragment\n    pinned {\n      personalUpdate\n      __typename\n    }\n    __typename\n  }\n  ... on Question {\n    ...UserPostFragment\n    __typename\n  }\n  ... on OfficialMessage {\n    ...LikeableFragment\n    ...CommentableFragment\n    ...MessageInfoFragment\n    ...RootMessageFragment\n    __typename\n  }\n  __typename\n}\n\nfragment EssentialFragment on MessageEssential {\n  id\n  type\n  content\n  shareCount\n  repostCount\n  createdAt\n  collected\n  pictures {\n    format\n    watermarkPicUrl\n    picUrl\n    thumbnailUrl\n    smallPicUrl\n    width\n    height\n    __typename\n  }\n  urlsInText {\n    url\n    originalUrl\n    title\n    __typename\n  }\n  __typename\n}\n\nfragment LikeableFragment on LikeableMessage {\n  liked\n  likeCount\n  __typename\n}\n\nfragment CommentableFragment on CommentableMessage {\n  commentCount\n  __typename\n}\n\nfragment RootMessageFragment on RootMessage {\n  topic {\n    id\n    content\n    __typename\n  }\n  __typename\n}\n\nfragment UserPostFragment on MessageUserPost {\n  readTrackInfo\n  user {\n    ...TinyUserFragment\n    __typename\n  }\n  __typename\n}\n\nfragment TinyUserFragment on UserInfo {\n  avatarImage {\n    thumbnailUrl\n    smallPicUrl\n    picUrl\n    __typename\n  }\n  isSponsor\n  username\n  screenName\n  briefIntro\n  __typename\n}\n\nfragment MessageInfoFragment on MessageInfo {\n  video {\n    title\n    type\n    image {\n      picUrl\n      __typename\n    }\n    __typename\n  }\n  linkInfo {\n    originalLinkUrl\n    linkUrl\n    title\n    pictureUrl\n    linkIcon\n    audio {\n      title\n      type\n      image {\n        thumbnailUrl\n        picUrl\n        __typename\n      }\n      author\n      __typename\n    }\n    video {\n      title\n      type\n      image {\n        picUrl\n        __typename\n      }\n      __typename\n    }\n    __typename\n  }\n  __typename\n}\n\nfragment RepostFragment on Repost {\n  target {\n    ...RepostTargetFragment\n    __typename\n  }\n  targetType\n  __typename\n}\n\nfragment RepostTargetFragment on RepostTarget {\n  ... on OriginalPost {\n    id\n    type\n    content\n    pictures {\n      thumbnailUrl\n      __typename\n    }\n    topic {\n      id\n      content\n      __typename\n    }\n    user {\n      ...TinyUserFragment\n      __typename\n    }\n    __typename\n  }\n  ... on Repost {\n    id\n    type\n    content\n    pictures {\n      thumbnailUrl\n      __typename\n    }\n    user {\n      ...TinyUserFragment\n      __typename\n    }\n    __typename\n  }\n  ... on Question {\n    id\n    type\n    content\n    pictures {\n      thumbnailUrl\n      __typename\n    }\n    user {\n      ...TinyUserFragment\n      __typename\n    }\n    __typename\n  }\n  ... on Answer {\n    id\n    type\n    content\n    pictures {\n      thumbnailUrl\n      __typename\n    }\n    user {\n      ...TinyUserFragment\n      __typename\n    }\n    __typename\n  }\n  ... on OfficialMessage {\n    id\n    type\n    content\n    pictures {\n      thumbnailUrl\n      __typename\n    }\n    __typename\n  }\n  ... on DeletedRepostTarget {\n    status\n    __typename\n  }\n  __typename\n}\n"
        }

        # 设置请求头

        res = requests.post('https://web-api.okjike.com/api/graphql', headers=self.headers, json=json,timeout=5).json()
        # 动态列表
        out = {}
        post_list = res['data']['userProfile']['feeds']['nodes']
        for i in post_list:
            info = {
                'date': i["createdAt"],
                'content': i["content"],
            }
            print(info)
            out[i['id']] = info

        last_id = post_list[-1]['id']

        while(1):
            sleep(2)
            __variables = {
            "username":__username,
            "loadMoreKey": {
                "lastId": last_id
                }
            }
            json['variables'] = __variables
            try:
                res = requests.post('https://web-api.okjike.com/api/graphql', headers=self.headers, json=json,timeout=5).json()
            except:
                break
            # 动态列表
            post_list = res['data']['userProfile']['feeds']['nodes']
            for i in post_list:
                info = {
                    'date': i["createdAt"],
                    'content': i["content"],
                }
                print(info)
                out[i['id']] = info
            if post_list[-1]['id'] == last_id:
                break
            last_id = post_list[-1]['id']

        print(len(out))
        return out

    def get_comments(self,post_id):
        json = {
        "operationName":"MessageComments",
        "variables":{
            "messageId":post_id,
            "messageType": "ORIGINAL_POST"
        },
        "query": "query MessageComments($messageType: MessageType!, $messageId: ID!, $loadMoreCommentKey: JSON) {\n  message(messageType: $messageType, id: $messageId) {\n    id\n    ... on OriginalPost {\n      comments(loadMoreKey: $loadMoreCommentKey) {\n        ...CommentConnectionFragment\n        __typename\n      }\n      __typename\n    }\n    ... on Repost {\n      comments(loadMoreKey: $loadMoreCommentKey) {\n        ...CommentConnectionFragment\n        __typename\n      }\n      __typename\n    }\n    __typename\n  }\n}\n\nfragment CommentConnectionFragment on CommentConnection {\n  pageInfo {\n    loadMoreKey\n    hasNextPage\n    __typename\n  }\n  nodes {\n    ...CommentFragment\n    __typename\n  }\n  __typename\n}\n\nfragment CommentFragment on Comment {\n  id\n  threadId\n  collapsed\n  collapsible\n  targetId\n  targetType\n  createdAt\n  level\n  content\n  user {\n    ...TinyUserFragment\n    __typename\n  }\n  urlsInText {\n    title\n    originalUrl\n    url\n    __typename\n  }\n  pictures {\n    format\n    picUrl\n    watermarkPicUrl\n    smallPicUrl\n    thumbnailUrl\n    width\n    height\n    __typename\n  }\n  likeCount\n  liked\n  replyCount\n  enablePictureComments\n  hotReplies {\n    ...InnerCommentFragment\n    __typename\n  }\n  __typename\n}\n\nfragment TinyUserFragment on UserInfo {\n  avatarImage {\n    thumbnailUrl\n    smallPicUrl\n    picUrl\n    __typename\n  }\n  isSponsor\n  username\n  screenName\n  briefIntro\n  __typename\n}\n\nfragment InnerCommentFragment on Comment {\n  id\n  threadId\n  createdAt\n  content\n  level\n  user {\n    ...TinyUserFragment\n    __typename\n  }\n  urlsInText {\n    title\n    originalUrl\n    url\n    __typename\n  }\n  pictures {\n    format\n    picUrl\n    thumbnailUrl\n    width\n    height\n    __typename\n  }\n  replyToComment {\n    user {\n      ...TinyUserFragment\n      __typename\n    }\n    __typename\n  }\n  __typename\n}\n"
        }
        res = requests.post('https://web-api.okjike.com/api/graphql', headers=self.headers, json=json,timeout=5).json()
        return res

memo = jike()
# posts = memo.show_all_post('6ba2b244-8c7c-479d-b840-dd6cc81de64a')
res = memo.get_comments('62d117e97423e1671598917e')

peer = []
comments = res['data']['message']['comments']['nodes']
for node in comments:
    content = node['content']
    teller_id = node['user']['username']
    peer.append([teller_id,content])

print(peer)

# print(build_excel(JSON_DATA=posts,excel_name_three='jike'+str(randint(1,100))))