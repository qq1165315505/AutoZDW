# 我的常用api
import os.path
import time

import pandas as pd
import base64
import json
import requests

def base64_api(img, uname = "qq1165315505", pwd = "qq111678", typeid = 1):
    with open(img, 'rb') as f:
        base64_data = base64.b64encode(f.read())
        b64 = base64_data.decode()
    data = {"username": uname, "password": pwd, "typeid": typeid, "image": b64}
    qheaders = {"User-Agent":
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36"}
    # 设置重连次数
    requests.adapters.DEFAULT_RETRIES = 5
    s = requests.session()
    # 设置重连活跃状态为False
    s.keep_alive = False
    res = requests.post("http://api.ttshitu.com/predict",
                        headers=qheaders,stream=True,verify=False,timeout=(10,10),json=data)
    result = json.loads(res.text)
    res.close()
    if result['success']:
        return result["data"]["result"]
    else:
        return result["message"]
    return ""

def loadCst(path):
    """
    下载待查客户数据
    :param path: 数据路径
    :return: 数据list
    """
    csts = pd.read_excel(path).drop_duplicates().values.tolist()
    cstsOut = []
    for cst in csts:
        cstsOut.append(cst[0])
    return cstsOut

def safeCst(csts,path):
    """保存数据"""
    df = pd.DataFrame(data=csts,columns=["公司名称"])
    df.to_excel(path,index=False)

def outData(path):
    """
    整理并导出数据
    :param path: 输出数据路径
    :return:
    """
    # excel文件名列表
    exLists = os.listdir(path)
    dfK = pd.DataFrame(columns=["承租人", "租赁公司", "起租日", "期限", "登记金额"],index=None)
    for exList in exLists:
        df = pd.read_excel(os.path.join(path, exList), sheet_name="融资租赁")
        print(df)
        dfMid = df[["填表人", "登记时间", "登记期限", "租赁财产价值","登记类型"]]
        dfMid["承租人"] = exList[:-4]
        dfMid.rename(columns={'承租人': '承租人',
                              '填表人': '租赁公司',
                              '登记时间': '起租日',
                              '登记期限': '期限',
                              '租赁财产价值': '登记金额',
                              '登记类型': '登记类型',})
        dfK = pd.concat([dfK,dfMid])
    # 删除重复值
    dfK = dfK.drop_duplicates()
    # 整理，删除变更登记
    dfOut = dfK[dfK["登记类型"] != "变更登记"]
    dfOut = dfOut[dfOut["登记类型"] != "展期登记"]
    for i in range(dfOut.shape[0]):
        if dfOut[i]["登记类型"] == "注销登记":
            dfOut = dfOut.drop(index=[i,i-1])
    return dfOut

def renameData(loadPath,company):
    """检测loadpath是否有文件进入，并改名"""
    dirList = os.listdir(loadPath)
    if not dirList:
        return False
    else:
        dirList = sorted(dirList,
                         key=lambda x: os.path.getmtime(os.path.join(loadPath, x)))
        # 得到最新文件夹
        latestDir = dirList[-1]
        # 最新文件的创建时间
        latestDirTime = os.path.getmtime(os.path.join(loadPath,latestDir))
        dtime = time.time() - latestDirTime
        if dtime < 124:
            os.rename(os.path.join(loadPath,latestDir),
                      os.path.join(loadPath,company+".xls"))

if __name__ == "__main__":
    # safeCst(loadCst("../data/cst.xls"),"../data/cst.xls")
    path = "../download0"
    company = "是我被修改1234"
    renameData(path,company)




