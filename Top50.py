import requests
from bs4 import BeautifulSoup
import openpyxl
import matplotlib.pyplot as plt
import pandas as pd

url = 'http://top.baidu.com/buzz?b=7&fr=topbuzz_b354'
h = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36",
}


def get_request():
    try:
        r = requests.get(url, headers=h)
        r.raise_for_status()  # 如果不是200，则引发HTTPError异常
        r.encoding = r.apparent_encoding  # 根据内容去确定编码格式
        return r.text
    except BaseException as e:
        print("出现异常：", e)
        return str(e)


def get_data():
    context = []
    print("开始爬虫")
    r = get_request()
    print("开始解析")
    soup = BeautifulSoup(r, 'html.parser')
    tr = soup.select('tr')
    for item in tr:
        new_item = []
        if item.select('span.num-top'):   # 发现前三名的和后面的标签不一样于是用判断语句分别获得
            num = item.select('span.num-top')[0].text
            book_name = item.select('a.list-title')[0].text
            if item.select('span.icon-rise'):   # 发现点击量有上升、下降、不变分别存在三个元素中，用多分支结构处理
                click = item.select('span.icon-rise')[0].text
            elif item.select('span.icon-fall'):
                click = item.select('span.icon-fall')[0].text
            else:
                click = item.select('span.icon-fair')[0].text
            new_item.append(num)
            new_item.append(book_name)
            new_item.append(click)
            context.append(new_item)
        elif item.select('span.num-normal'):
            num = item.select('span.num-normal')[0].text
            book_name = item.select('a.list-title')[0].text
            if item.select('span.icon-rise'):
                click = item.select('span.icon-rise')[0].text
            elif item.select('span.icon-fall'):
                click = item.select('span.icon-fall')[0].text
            else:
                click = item.select('span.icon-fair')[0].text
            new_item.append(num)
            new_item.append(book_name)
            new_item.append(click)
            context.append(new_item)
    return context


def writefile(file_name, content_str):   # 写入txt文件
    with open(file_name, "w", encoding='utf-8', ) as f:
        f.write('{:<12}{:<12}{:<12}\n'.format("排名", "书名", "点击量"))
        for content in content_str:  # 外循环获得每本书的信息
            for i in content:  # 内循环获得每本书的属性
                f.write('{:<12}'.format(i))
            f.write('\n')
        f.close


def write_excel(file_name, list_content):
    wb = openpyxl.Workbook()  # 新建Excel工作簿
    st = wb.active
    st['A1'] = "百度小说top50"  # 修改为自己的标题
    second_row = ["排名", "书名", "点击量"]  # 根据实际情况写属性
    st.append(second_row)
    st.merge_cells("A1:B1")  # 根据实际情况合并标题单元格
    for row in list_content:
        st.append(row)
    wb.save(file_name)  # 新工作簿的名称


def show_bar(list_content):  # 制图
    x = []
    y = []
    for content in list_content:
        x.append(content[1])
        y.append(int(content[2]))
    plt.rcParams['font.sans-serif'] = ['KaiTi']
    plt.title("百度小说TOP50点击量")
    plt.xlabel("书名")  # x轴标签
    plt.ylabel("点击量")  # y轴标签
    plt.bar(x, y,  width=0.5, linewidth=5.0)
    plt.xticks(rotation=60)
    # plt.legend(labels=["评分"],loc="best")
    plt.savefig('top50.png')
    plt.show()



def write_csv(file_name, content_str):
    new_content = []
    for i in content_str:
        new_content.append(i[1:])
    title = ["书名", "点击量"]
    top = pd.DataFrame(columns=title, data=new_content)
    top.to_csv(file_name)


if __name__ == '__main__':
    txt = './小说TOP榜.txt'
    excel = './小说TOP榜.xlsx'
    csv = './小说TOP榜.csv'
    contents = get_data()
    writefile(txt, contents)
    write_excel(excel, contents)
    write_csv(csv, contents)
    show_bar(contents)