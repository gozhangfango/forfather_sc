import xml.etree.ElementTree as ET  # 这个模块负责解析xml配置文件
import PIL.ImageGrab                         # 这个模块负责屏幕截图
import pytesseract
import matplotlib.pyplot as plt # plt 用于显示图片
import matplotlib.image as mpimg # mpimg 用于读取图片

class RectItem:
    def __init__(self):
        self.name = ''  # 名字
        self.topleftwidth = 0  # 左上点横坐标
        self.topleftheight = 0  # 左上点纵坐标
        self.bottomrightwidth = 0  # 右下点横坐标
        self.bottomrightheight = 0  # 右下点纵坐标
        self.count = 0

def main():
    # 第一部分，读取配置文件
    print('分析配置文件')
    tree = ET.ElementTree(file='config.xml')
    root = tree.getroot()
    rectlist = []
    displaypar = 1
    step = 10
    displaypar = int(tree.find("display-par").text)
    step = tree.find("step").text
    for elem in tree.iter(tag='rect'):
        tempitem = RectItem()
        tempitem.name = elem.get("name")
        tempitem.topleftwidth = int(elem.find("top-left-dot").get("width")) * displaypar
        tempitem.topleftheight = int(elem.find("top-left-dot").get("height")) * displaypar
        tempitem.bottomrightwidth = int(elem.find("bottom-right-dot").get("width")) * displaypar
        tempitem.bottomrightheight = int(elem.find("bottom-right-dot").get("height")) * displaypar
        rectlist.append(tempitem)

    for index in range(len(rectlist)):
        im = PIL.ImageGrab.grab((rectlist[index].topleftwidth, rectlist[index].topleftheight,
                                 rectlist[index].bottomrightwidth, rectlist[index].bottomrightheight))
        # im = im.resize((im.size[0] * 18, im.size[1] * 18), PIL.Image.ANTIALIAS)
        addr = r'testimage/' + rectlist[index].name + ".png"
        im.save(addr, 'png')
        tmpstr = pytesseract.image_to_string(im, lang="num", config="-psm 8")
        try:
            result = float(tmpstr)
            tmp_print = f"点位{index+1}，识别成功，结果是：{result}"
            print(tmp_print)
        except Exception as e:
            tmp_print = f"点位{index+1}，识别不成功，结果是：{tmpstr}"
            print(tmp_print)



    fig = plt.figure(figsize=(12,6))
    x = 0.01
    y = 0.85
    for rect in rectlist:
        a = plt.axes([x, y, .1, .1])
        lena = mpimg.imread('testimage/' + rect.name + '.png')
        plt.imshow(lena)  # 显示图片
        plt.title(rect.name)
        a.spines['right'].set_color('none')
        a.spines['top'].set_color('none')
        a.spines['bottom'].set_color('none')
        a.spines['left'].set_color('none')
        a.set_xticks([])

        a.set_yticks([])
        x = x + 0.11
        if x > 0.8:
            x = 0.01
            y = y - 0.11

    plt.show()




if __name__ == '__main__':
    main()
    rectlist = []