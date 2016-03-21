import requests
from html.parser import HTMLParser

class NextPageParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tmp = None
        self.nexta = None

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            self.tmp = attrs

    def handle_endtag(self, tag):
        pass

    def handle_data(self, data):
        if '下一页' in data:
            self.nexta = self.tmp

def get_next_page_url(text):
    parser = NextPageParser()
    parser.feed(text)
    taga = parser.nexta
    if taga is None:
        return None
    for i in taga:
        if i[0] =='href':
            return i[1]
    return None

class MyHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.data = list()
        self.tmp = ''
        self.is_in_a_tag = False
        self.current_tag = None

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            self.tmp = ''
            self.is_in_a_tag = True
        for k, v in attrs:
            if k == 'class' and v == 'c-showurl':
                self.tmp = ''
                self.current_tag = tag

    def handle_endtag(self, tag):
        if tag =='a':
            self.data.append(self.tmp)
            self.is_in_a_tag = False

        if tag == self.current_tag:
            self.data.append(self.tmp)
            self.current_tag = None

    def handle_data(self, data):
        if self.current_tag is not None:
            self.tmp = self.tmp + data
        if data == '百度快照':
            self.tmp = self.tmp + data

    def getdata(self):
        return self.data

def getranklist(text):
    parser = MyHTMLParser()
    parser.feed(text)
    t = parser.getdata()
    stack = list()
    rank = list()
    for i in t:
        stack.append(i)
        if i == '百度快照':
            shapshot = stack.pop().replace(u'\xa0', u' ')
            dash = stack.pop().replace(u'\xa0', u' ')
            addr = stack.pop().replace(u'\xa0', u' ')
            if addr.replace(' ', '') == '':
                addr = stack.pop().replace(u'\xa0', u' ')
            rank.append(addr)

    return rank

class Rank(object):
    def __init__(self, keyword):
        self.keyword = keyword
        self.init_url = '/s?wd={0}'.format(self.keyword)
        self.session = requests.Session()

    def gethtml(self, url = None):
        baidu_url = 'http://www.baidu.com{0}'
        if url is None:
            baidu_url = baidu_url.format(self.init_url)
        else:
            baidu_url = baidu_url.format(url)

        resp = self.session.get(baidu_url)

        if resp.status_code == 200:
            return resp.text
        else:
            return None

    def getRank(self, num = 5):
        next_url = None
        rank = list()
        while len(rank) < num:
            htmltext = self.gethtml(next_url)
            [rank.append(i) for i in getranklist(htmltext)]
            next_url = get_next_page_url(htmltext)
            print(next_url)
            if next_url is None:
                print('next url is None')
                break

        return rank[:num]

if __name__ == "__main__":
    x = Rank('诺基亚')
    l = x.getRank(20)
    for i in l:
        print(i)
