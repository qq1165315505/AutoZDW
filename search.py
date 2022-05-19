# 中登网登录并查询
import os
import threading
import time


from script.myUtils import *
from data.userInfo import getUserInfo
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

class ZDWSearch():
    def __init__(self,userInfo,lock,csts,key):
        self.csts = csts
        self.lock = lock
        self.userInfo = userInfo
        self.loadPath = os.path.join(r"D:\coderlife\auto_zdw",
                                     "download"+str(key))  # 下载地址
        
        self.drv = self.initWebDriver()  # 浏览器实例化
        # 客户名单地址
        self.cstPath = "data/cst.xls"

    # 初始化工作************************************************************

    def initWebDriver(self):
        """
        初始化浏览器
        :param tmp_path: 文件下载路径
        :return:
        """
        with self.lock:
            # 创建selenium启动配置参数类
            option = webdriver.ChromeOptions()
            # 配置driver参数：
            # 1、添加启动参数
            # 浏览器UA信息
            UA = 'user-agent="Mozilla/5.0 (Windows NT 10.0; WOW64)' \
                 ' AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36"'
            option.add_argument(UA)
            # 最大化浏览器
            option.add_argument("start-maximized")
            # 规避谷歌浏览器bug
            option.add_argument("--disable-gpu")
            # # 浏览器不提供可视化页面. linux下如果系统不支持可视化不加这条会启动失败
            # option.add_argument('--headless')
            # 以最高权限运行
            option.add_argument('--no-sandbox')
            # 设置下载路径
            prefs = {'profile.default_content_settings.popups': 0,  # 防止保存弹窗
                     'download.default_directory': self.loadPath,  # 设置默认下载路径
                     "profile.default_content_setting_values.automatic_downloads": 1,  # 允许多文件下载
                     "profile.default_content_setting_values": {"notifications": 2},  # 禁用弹窗
                     "credentials_enable_service":False,
                     "profile.password_manager_enabled": False,  #关掉密码弹窗
            }
            option.add_experimental_option('prefs', prefs)

            # 修改windows.navigator.webdriver，防机器人识别机制，selenium自动登陆判别机制
            option.add_experimental_option('excludeSwitches', ['enable-automation'])

            # 选择谷歌浏览器
            s = Service(executable_path=r'chromedriver.exe')
            # 实例化
            driver=webdriver.Chrome(service = s,options=option)
            # 返回webDriver
            return driver

    # 基础操作
    def input(self,by,value,text):
        """
        输入数据
        :param loc: ID 定位
        :param text: 输入数据
        :return:
        """
        with self.lock:
            elem = self.drv.find_element(by,value)
            elem.click()
            elem.clear()
            elem.send_keys(text)

    def findElement(self,by,value):
        with self.lock:
            return self.drv.find_element(by,value)

    def getCheckImgNum(self,by,value):
        """
        获得验证麻截图并保存在checkImg文件夹
        :param by:
        :param value:
        :return:
        """
        with self.lock:
            imgPath = "checkImg/checkImg.png"
            ele = self.drv.find_element(by,value)
            ele.click()
            ele.screenshot(imgPath)
            veriCode = base64_api(imgPath)
            os.remove(imgPath)
            if veriCode.isdigit():
                return int(veriCode)
            else:
                return False

    def waitElem(self,by,value,t = 300):
        """
        等待元素出现
        :param by: 查找方式
        :param value: 元素标识符
        :param t: 等待时间
        :return: 等待元素
        """
        res = WebDriverWait(self.drv,t,1).\
            until(EC.visibility_of_element_located((by,value)))
        return res

    def switchToDefault(self):
        """转到最外页面"""
        self.drv.switch_to.default_content()

    def close(self):
        """关闭浏览器"""
        self.drv.close()

    def isAleart(self):
        """判断是否有弹出警告出现"""
        time.sleep(1)
        with self.lock:
            try:
                alert = EC.alert_is_present()(self.drv)
                print("警告，%s出现alert窗口，内容为%s"%
                      (threading.current_thread().getName(),alert.text))
                alert.accept()
                return True
            except:
                return False

    # 登录操作************************************************************
    def login(self):
        """登录中登网，并输入账号密码，验证登录信息"""
        self.drv.get("https://www.zhongdengwang.org.cn/")
        # 点击用户登录
        self.findElement(By.XPATH, '//*[@id="login"]/div/div[1]/a[1]').click()
        # 输入用户名
        self.input(By.ID,"usercode",self.userInfo["user"])
        # 设置当前密码库指引
        pswIndex = 0
        # 密码错误次数
        pswInputErrorTimes = 0
        # 登录流程
        for i in range(10):
            # 输入密码
            self.input(By.ID, "password", self.userInfo["psw"][pswIndex])
            # 获得验证码并输入
            for j in range(10):
                veriCode = self.getCheckImgNum(By.ID,"checkImg")
                if veriCode:
                    self.input(By.ID,"verifyCode",veriCode)
                    break
            else:
                raise Exception("验证码错误，请检测")
            # 点击登录
            self.findElement(By.ID, "btn_login").click()
            """
            登录情况反馈*****************************************************
            error:
            1、登录名或密码错误，若连续5次错误将锁定
            2、校验码错误!
            right:
            手机验证码阶段
            """
            # 错误情况：1、登录名或密码错误，若连续5次错误将锁定
            eleText = self.waitElem(By.ID,"bsmodal").text
            if "验证手机号" in eleText:
                # 点击获得验证码
                sendCodeFlag = True
                sendCodeEle = self.findElement(By.ID, "sendCode")
                while sendCodeFlag:
                    sendCodeEle.click()
                    if sendCodeEle.text != "获取动态码":
                        sendCodeFlag = False

                for k in range(10):
                    # 输入手机验证码
                    with self.lock:
                        telNum = input("请输入%s的验证码：" % self.userInfo["tel"])
                    self.input(By.ID, "dynamicCode", telNum)
                    self.findElement(By.NAME,"okbtn").click()
                    # 判断是否有验证码错误
                    if self.isAleart():
                        # 有弹窗，重新输入验证码
                        continue
                    else:
                        # 无弹窗，进入登录页面
                        print("登录成功！")
                        return True
            else:
                """
                错误情况:
                1、登录名或密码错误，若连续5次错误将锁定
                2、校验码错误!
                3、其他为识别
                """
                if "校验码错误!" in eleText:
                    self.findElement(By.NAME,"okbtn").click()
                    # 回到循环开头，重新输入密码，验证码
                    continue
                elif "登录名或密码错误，若连续5次错误将锁定" in eleText:
                    if pswInputErrorTimes <=3:
                        self.findElement(By.NAME,"okbtn").click()
                        pswInputErrorTimes += 1
                        if pswIndex == 0:
                            pswIndex = 1
                        else:
                            pswIndex = 0
                        continue
                    else:
                        raise Exception("密码库错误！请检测密码框！")
                else:
                    raise Exception("在登录账号发现其他错误！")

    # 查询操作************************************************************
    def search(self,company):
        """
        查询中登网数据并下载
        :param company: 公司名称
        :return: True or False
        """
        # 返回父frame
        self.switchToDefault()
        time.sleep(1)
        # 进入查询界面
        menudivEle = self.waitElem(By.XPATH,"//*[@id='menudiv']/ul/li[2]/a")
        if not menudivEle.get_attribute("class"):
            menudivEle.click()
        self.waitElem(By.XPATH, "//*[@id='menudiv']/ul/li[2]/ul/li[1]/a").click()
        # 输入验证码
        for i in range(10):
            # 转向查询iframe的html
            self.drv.switch_to.frame(self.findElement(By.ID, "first-iframe"))
            # 担保人类型
            self.waitElem(By.XPATH, "//*[@onclick='checkOrgan();']").click()
            # 输入公司名称
            self.input(By.ID, "organNameInput", company)
            # 获得验证码并输入
            for j in range(10):
                veriCode = self.getCheckImgNum(By.ID, "checkImg")
                if veriCode:
                    self.input(By.ID, "verifyCode", veriCode)
                    break
            else:
                raise Exception("验证码错误，请检测")
            # 查询证明选项
            self.findElement(By.ID, "isNeedProveNo").click()
            # 查询
            self.waitElem(By.XPATH,"//*[@value='查 询']").click()
            # 错误情况
            try:
                """
                错误情况：
                1、请填写机构名称
                2、请填写校验码
                3、请选择是否需要查询证明
                4、验证码错误
                """
                self.switchToDefault()
                self.waitElem(By.ID, "bsmodal", t=3)
                print("有弹窗！")
                self.waitElem(By.NAME, "okbtn").click()
                continue
            except:
                print("无弹窗")
                pass
            # 下载

            # 正确情况
            # 查看是否查询到数据，其中searchNum为查询数据数目
            try:
                # 转向查询iframe的html
                self.drv.switch_to.frame(self.findElement(By.ID, "first-iframe"))
                # 拿数据
                parentNum = int(self.waitElem(By.ID, "registCount").text)
                # 若有数据
                if parentNum > 0:
                    qizhongText = self.findElement(By.XPATH, "//span[@id='qizhong']/..").text
                    print(threading.current_thread().getName(),
                          company,"有",parentNum,"笔登记！主要明细如下：\n",qizhongText)
                    if "融资租赁" in qizhongText:
                        searchNum = parentNum
                    else:
                        searchNum = 0
                else:
                    searchNum = 0
            except:
                searchNum = 0
            # searchNum为查询数据数目大于0说明有登记情况，等待下载
            print("查询结果数据量：",searchNum)
            if searchNum > 0:
                print("下载")
                self.waitElem(By.ID,"downloadDetailLink").click()
                # 等待查询时间
                for i in range(124):
                    time.sleep(1)
                    if i % 5 == 0:
                        print("%s剩余%d秒"%(threading.current_thread().getName(),124-i))
                    if i == 100:
                        self.renameData(company)
                print("%s查询的%s下载完毕"%(threading.current_thread().getName(),company))
                return True
            # 无登记情况！
            else:
                print("%s查询的%s无中等数据"%(threading.current_thread().getName(),company))
                return False

    """
    数据层面操作*****************************************************
    """
    def getCompanyName(self):
        """
        弹出公司名字
        :return: company
        """
        with self.lock:
            try:
                company = self.csts.pop()
                print("目前还剩",len(self.csts),"个公司待查")
                return company
            except:
                return False

    def safeCsts(self):
        """
        保存客户名单
        :return:
        """
        with self.lock:
            safeCst(self.csts,self.cstPath)

    def renameData(self,company):
        """检测loadpath是否有文件进入，并改名"""
        print(self.loadPath,"路径正在改名字")
        dirList = os.listdir(self.loadPath)
        if not dirList:
            return False
        else:
            dirList = sorted(dirList,
                             key=lambda x: os.path.getmtime(os.path.join(self.loadPath, x)))
            # 得到最新文件夹
            latestDir = dirList[-1]
            # 最新文件的创建时间
            latestDirTime = os.path.getmtime(os.path.join(self.loadPath,latestDir))
            dtime = time.time() - latestDirTime
            if dtime < 125:
                os.rename(os.path.join(self.loadPath,latestDir),
                          os.path.join(self.loadPath,company+".xls"))

    """
    主执行层*****************************************************
        """
    def run(self):
        if self.login():
            cstIndex = 0
            FinishFlag = False
            while not FinishFlag:
                company = self.getCompanyName()
                if company:
                    self.search(company)
                    cstIndex += 1
                    if cstIndex %5 == 0:
                        print("每查询5次保存一次数据记录")
                        self.safeCsts()
                else:
                    print("数据查询完毕！")
                    FinishFlag = False
        else:
            raise Exception("登录出错！")


if __name__ == "__main__":
    # 登录账号及查询数据下载
    userInfos = getUserInfo()
    csts = loadCst("data/cst.xls")
    # 线程锁
    lock = threading.Lock()
    # 线程与程序
    opts = []
    tds = []
    for key,userInfo in userInfos.items():
        if key < 5:
            opts.append(ZDWSearch(userInfo,lock,csts,key))
    for index,opt in enumerate(opts):
        td = threading.Thread(target=opt.run,name=index)
        tds.append(td)
        td.start()
    for td in tds:
        td.join()
    if csts == []:
        outData("finalData.xls")




