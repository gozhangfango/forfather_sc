import xml.etree.ElementTree as ET  # 这个模块负责解析xml配置文件
import PIL.ImageGrab                         # 这个模块负责屏幕截图
from apscheduler.schedulers.blocking import BlockingScheduler
import uuid
import pytesseract

class RectItem:
    def __init__(self):
        self.name = ''  # 名字
        self.topleftwidth = 0  # 左上点横坐标
        self.topleftheight = 0  # 左上点纵坐标
        self.bottomrightwidth = 0  # 右下点横坐标
        self.bottomrightheight = 0  # 右下点纵坐标
        self.count = 0

if __name__ == '__main__':
    #第一部分，读取配置文件
    print('分析配置文件')
    tree = ET.ElementTree(file='config.xml')
    root = tree.getroot()
    rectlist = []
    displaypar = 1
    step =10
    displaypar = int(tree.find("display-par").text)
    step = tree.find("step").text
    for elem in tree.iter(tag='rect'):
        tempitem = RectItem()
        tempitem.name = elem.get("name")
        tempitem.topleftwidth = int(elem.find("top-left-dot").get("width"))*displaypar
        tempitem.topleftheight = int(elem.find("top-left-dot").get("height"))*displaypar
        tempitem.bottomrightwidth = int(elem.find("bottom-right-dot").get("width"))*displaypar
        tempitem.bottomrightheight = int(elem.find("bottom-right-dot").get("height"))*displaypar
        rectlist.append(tempitem)


    #第二部分，开始抓图
    def catchjob():
        for rect in rectlist:
            im = PIL.ImageGrab.grab((rect.topleftwidth, rect.topleftheight, rect.bottomrightwidth, rect.bottomrightheight))
            #im = im.resize((im.size[0] * 18, im.size[1] * 18), PIL.Image.ANTIALIAS)
            tmpstr = pytesseract.image_to_string(im, lang="num")
            try:
                result = float(tmpstr)
                addr = r'trainimage/' + tmpstr + "----" + str(uuid.uuid1()) + ".png"
                im.save(addr, 'png')
            except Exception as e:
                addr = r'wrongimage/' + tmpstr + "----" + str(uuid.uuid1()) + ".png"
                im.save(addr, 'png')


    par_str = r'*/' + step
    sched = BlockingScheduler()
    sched.add_job(catchjob, 'cron', second=par_str, minute='*', hour='*')
    print("开始抓取")
    print("程序现在每10秒截取一下图片，可以识别的，放在trainimage文件夹，不能识别的，放在wrongimage文件夹")
    sched.start()


