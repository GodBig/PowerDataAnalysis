import time
import requests
from datetime import datetime
import xlwt
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pytz
from PIL import Image


def make_bmp(filename):
    path = ".//log//file//"
    try:
        img = Image.open(path + filename + ".png")
        r, g, b, a = img.split()
        img = Image.merge("RGB", (r, g, b))
        img.save(path + filename + ".bmp")
        return 0
    except Exception as err:
        print("图片转换错误: ", err)
        return 1


def login():
    url = "http://172.24.141.73:8090/Dms/login.do"
    headers = {
        "Host": "172.24.141.73:8090",
        "Connection": "keep-alive",
        "Accept": "application/json, text/plain, */*",
        "Origin": "http://172.24.141.73:8090",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36",
        "Content-Type": "application/json;charset=UTF-8",
        "Referer": "http://172.24.141.73:8090/Dms/login.html",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cookie": ""
    }
    params = {"username": "aGF5al9qag==", "password": "RkFGRjJFMkQyN0I0MUJERjg3NkM4N0RDRUVGQzc0NjhlbmhoYzNGM01USWg=",
              "redirectUrl": "ajax"}
    try:
        res = requests.post(url=url, json=params, headers=headers, timeout=15)
        if res.json()["status"] == "success":
            cookie = requests.utils.dict_from_cookiejar(res.cookies)
            cookie_str = "JSESSIONID=" + cookie["JSESSIONID"] + "; secretCookieKey=" + cookie["secretCookieKey"]
        return cookie_str
    except Exception as err:
        print("重新运行,运行错误:", err)
        time.sleep(1)


def find_rcu(cookie):
    url = "http://172.24.141.73:8090/Dms/dms/common/model/loadStationGraphs.do"
    headers = {
        "Host": "172.24.141.73:8090",
        "Connection": "keep-alive",
        "Accept": "application/json, text/plain, */*",
        "Origin": "http://172.24.141.73:8090",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36",
        "Content-Type": "application/json;charset=UTF-8",
        "Referer": "http://172.24.141.73:8090/Dms/dashboard.do",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cookie": cookie
    }
    params = {"orgId": 8, "orgLevel": 2, "pageIndex": 1, "pageSize": -1}
    # 调整此参数可以监控全淮安
    break_flag = 0
    while True:
        try:
            res = requests.post(url=url, json=params, headers=headers, timeout=15)
            id_dict = res.json()
            return id_dict["data"]
        except Exception as err:
            if break_flag >= 20:
                break
            else:
                break_flag += 1
            print("重新运行,运行错误:", err)
            time.sleep(1)


def find_line(cookie, stationId):
    url = "http://172.24.141.73:8090/Dms/dms/common/model/loadFeederGraphs.do"
    headers = {
        "Host": "172.24.141.73:8090",
        "Connection": "keep-alive",
        "Accept": "application/json, text/plain, */*",
        "Origin": "http://172.24.141.73:8090",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36",
        "Content-Type": "application/json;charset=UTF-8",
        "Referer": "http://172.24.141.73:8090/Dms/dashboard.do",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cookie": cookie
    }
    params = {"stationId": stationId, "pageIndex": 1, "pageSize": -1}
    break_flag = 0
    while True:
        try:
            res = requests.post(url=url, json=params, headers=headers, timeout=15)
            id_dict = res.json()
            return id_dict["data"]
        except Exception as err:
            if break_flag >= 20:
                break
            else:
                break_flag += 1
            print("重新运行,运行错误:", err)
            time.sleep(1)


def find_feeder(cookie, feederId):
    url = "http://172.24.141.73:8090/Dms/dms/common/model/loadRtuList.do"
    headers = {
        "Host": "172.24.141.73:8090",
        "Connection": "keep-alive",
        "Accept": "application/json, text/plain, */*",
        "Origin": "http://172.24.141.73:8090",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36",
        "Content-Type": "application/json;charset=UTF-8",
        "Referer": "http://172.24.141.73:8090/Dms/dashboard.do",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cookie": cookie
    }
    params = {"feederId": feederId, "pageIndex": 1, "pageSize": -1}
    break_flag = 0
    while True:
        try:
            res = requests.post(url=url, json=params, headers=headers, timeout=15)
            id_dict = res.json()
            return id_dict["data"]
        except Exception as err:
            if break_flag >= 20:
                break
            else:
                break_flag += 1
            print("重新运行,运行错误:", err)
            time.sleep(1)


def find_rtu(cookie, rtuId):
    url = "http://172.24.141.73:8090/Dms/dms/common/model/loadRtuPoints.do"
    headers = {
        "Host": "172.24.141.73:8090",
        "Connection": "keep-alive",
        "Accept": "application/json, text/plain, */*",
        "Origin": "http://172.24.141.73:8090",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36",
        "Content-Type": "application/json;charset=UTF-8",
        "Referer": "http://172.24.141.73:8090/Dms/dashboard.do",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cookie": cookie
    }
    params = {"rtuId": rtuId, "pntTable": "pnt_ana"}
    break_flag = 0
    while True:
        try:
            res = requests.post(url=url, json=params, headers=headers, timeout=15)
            id_dict = res.json()
            return id_dict["data"]
        except Exception as err:
            if break_flag >= 20:
                break
            else:
                break_flag += 1
            print("重新运行,运行错误:", err)
            time.sleep(1)


def get_data(cookie, id, name):
    url = "http://172.24.141.73:8090/Dms/dms/common/dataAccess/loadHistoryChartData.do"
    headers = {
        "Host": "172.24.141.73:8090",
        "Connection": "keep-alive",
        "Accept": "application/json, text/plain, */*",
        "Origin": "http://172.24.141.73:8090",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36",
        "Content-Type": "application/json;charset=UTF-8",
        "Referer": "http://172.24.141.73:8090/Dms/dashboard.do",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Cookie": cookie
    }
    # 分析起始时间
    start_time = int(time.time() * 1000) - 3600 * 1000 * 24 * 5
    end_time = int(time.time() * 1000)
    params = {"startTime": start_time, "endTime": end_time, "timeSpace": 300, "pointList": [
        {"tagName": "scada.pnt_ana." + str(id) + ".value",
         "description": str(name)}]}
    break_flag = 0
    while True:
        try:
            res = requests.post(url=url, json=params, headers=headers, timeout=15).json()
            return res
        except Exception as err:
            if break_flag >= 20:
                break
            else:
                break_flag += 1
            print("重新运行,运行错误:", err)
            time.sleep(1)


def stamptodate(time_stamp):
    # 先变成时间数组
    ufc_8 = str(datetime.fromtimestamp(time_stamp, pytz.timezone('Asia/Shanghai')).strftime('%Y-%m-%d-%H:%M:%S'))
    # 转换成新的时间格式(2016-05-05 20:28:54)
    return ufc_8


def stamptofileneme(time_stamp):
    # 先变成时间数组
    ufc_8 = str(datetime.fromtimestamp(time_stamp, pytz.timezone('Asia/Shanghai')).strftime('%Y-%m-%d'))
    # 转换成新的时间格式(2016-05-05 20:28:54)
    return ufc_8


def make_pic(data_list, pic_name):
    key_list = []
    for key in data_list[0].keys():
        key_list.append(key)
    key_1 = key_list[0]
    key_2 = key_list[1]
    time_list = []
    row_list = []
    for data in data_list:
        time_list.append(data[key_1])
        row_list.append(data[key_2])
    fig = plt.figure(pic_name, figsize=(12, 8))
    ax = fig.add_subplot()
    ax.plot(time_list,  # x轴数据
            row_list,  # y轴数据
            color='blue',
            linestyle='-',
            linewidth=1,
            markersize=1,
            markeredgecolor='black',
            markerfacecolor='brown')
    tick_spacing = 100
    ax.xaxis.set_major_locator(ticker.MultipleLocator(tick_spacing))
    plt.xticks(rotation=10)
    plt.xlabel(time_list[0] + "===>" + time_list[-1])
    plt.savefig(".//log//file//" + pic_name + ".png")
    plt.clf()
    plt.cla()
    make_bmp(pic_name)


def make_excel(collection, filename):
    file = xlwt.Workbook()
    sheet = file.add_sheet("淮安配电自动化平台统计")
    # 指定存储路径，如果当前路径存在同名文件，会覆盖掉同名文件
    style = xlwt.XFStyle()
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中
    style.alignment = al
    font = xlwt.Font()
    font.height = 20 * 14
    style.font = font
    # 创建一个样式对象，初始化样式
    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 20 * 40
    # 高度设置完
    sheet.col(0).width = 256 * 20
    sheet.col(1).width = 256 * 40
    sheet.col(2).width = 256 * 80
    sheet.col(3).width = 256 * 60
    sheet.col(4).width = 256 * 40
    sheet.col(5).width = 256 * 60
    sheet.write(0, 0, "变电站", style)
    sheet.write(0, 1, "供电线路", style)
    sheet.write(0, 2, "所属开关", style)
    sheet.write(0, 3, "无线终端", style)
    sheet.write(0, 4, "缺陷时间", style)
    sheet.write(0, 5, "数据截图", style)
    row = 1
    for document in collection:
        sheet.row(row).height_mismatch = True
        sheet.row(row).height = 20 * 200
        sheet.write(row, 0, document["rcu_name"], style)
        sheet.write(row, 1, document["line_name"], style)
        sheet.write(row, 2, document["feeder_dict"], style)
        sheet.write(row, 3, document["rtu_dict"], style)
        sheet.write(row, 4, document["err_time"], style)
        sheet.write(row, 5, "详情见淮安配电自动化平台", style)
        sheet.insert_bitmap(".//log//file//" + document["pic_name"] + ".bmp", row, 5, x=0.3, y=0.3, scale_x=0.34,
                            scale_y=0.03)
        row = row + 1
    file.save(".//log//" + filename + ".xls")


def main():
    err_list = []
    print(stamptodate(time.time()), "正在运行:")
    print(stamptodate(time.time())[11:15])
    cookie_str = login()
    rcu_list = find_rcu(cookie_str)
    ex_rcu_dict = ""
    ex_line_dict = ""
    ex_feeder_dict = ""
    ex_rtu_dict = ""
    for rcu_dict in rcu_list:
        print("正在查看:", rcu_dict["description"])
        line_list = find_line(cookie_str, int(rcu_dict["id"]))
        for line_dict in line_list:
            feeder_list = find_feeder(cookie_str, line_dict["id"])
            for feeder_dict in feeder_list:
                rtu_list = find_rtu(cookie_str, feeder_dict["id"])
                for rtu_dict in rtu_list:
                    for chance in range(2):
                        try:
                            if rtu_dict["description"]:
                                descrip_list = rtu_dict["description"].split(".")
                                if descrip_list[-1] == "零序电流":
                                    data_dict = get_data(cookie_str, rtu_dict["id"], rtu_dict["description"])
                                    if len(data_dict["data"]["columns"]) > 1:
                                        data_key = data_dict["data"]["columns"][-1]
                                        data_list = []
                                        for row_dict in data_dict["data"]["rows"]:
                                            data_list.append(float(row_dict[data_key]))
                                        if len(data_list) > 1:
                                            sort_list = sorted(data_list)
                                            mid_data = sort_list[int(len(sort_list) / 2)]
                                            max_data = sort_list[-1]
                                            if max_data - mid_data >= 1:
                                                if ex_rcu_dict == rcu_dict["description"] and \
                                                        ex_line_dict == line_dict["description"] and \
                                                        ex_feeder_dict == feeder_dict["description"] and \
                                                        ex_rtu_dict == rtu_dict["description"]:
                                                    pass
                                                else:
                                                    print("--------------------------------------------------------")
                                                    print("查看web端==>")
                                                    print("\t故障变电站:", rcu_dict["description"])
                                                    print("\t故障线路:", line_dict["description"])
                                                    print("\t故障节点:", feeder_dict["description"])
                                                    print("\t故障数据:", rtu_dict["description"])
                                                    err_dict = {"rcu_name": rcu_dict["description"],
                                                                "line_name": line_dict["description"],
                                                                "feeder_dict": feeder_dict["description"],
                                                                "rtu_dict": rtu_dict["description"],
                                                                "err_time": data_dict["data"]["rows"][-1]["时间"],
                                                                "pic_name": rtu_dict["description"].replace(" ",
                                                                                                            "").replace(
                                                                    ".", "-")}
                                                    err_list.append(err_dict)
                                                    make_pic(data_dict["data"]["rows"], err_dict["pic_name"])
                                                    ex_rcu_dict = rcu_dict["description"]
                                                    ex_line_dict = line_dict["description"]
                                                    ex_feeder_dict = feeder_dict["description"]
                                                    ex_rtu_dict = rtu_dict["description"]
                        except Exception as err:
                            print("数据出现问题:", err)
    make_excel(err_list, stamptofileneme(time.time()) + "-配电平台报告")
    print(stamptodate(time.time()), "运行结束")


def keep_run():
    while True:
        try:
            if stamptodate(time.time())[11:15] == "08:1":
                main()
                time.sleep(3600 * 24 - 600)
            elif stamptodate(time.time())[11:14] == "07:" or stamptodate(time.time())[11:14] == "08:":
                time.sleep(1)
            else:
                time.sleep(3600)
        except Exception as err:
            print("数据出现问题:", err)


if __name__ == '__main__':
    main()
    # print(login())
