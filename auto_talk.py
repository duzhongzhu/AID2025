import pandas as pd
"""
pandas 是一个 Python 的数据分析库，用于数据处理和分析。
它提供了高效的数据结构（如 DataFrame 和 Series）和多种操作数据的功能，
适用于各种格式的数据（如 CSV、Excel、SQL 数据库等）。

pandas 的核心数据结构是：

Series：一种类似于一维数组的对象，带有标签（索引）。
DataFrame：一个二维表格，类似于 Excel 表格或 SQL 数据表，包含行和列。
"""
import numpy as np

from uiautomation import WindowControl,MenuControl

#绑定微信主窗口
wx=WindowControl(Name='微信',
                 #searchDepth=1
 )
print(wx)
#切换窗口
wx.SwitchToThisWindow()
#寻找会话控件绑定
hw=wx.ListControl(Name='会话')
print('寻找会话控件绑定',hw)
#通过pd读取数据
df = pd.read_csv('回复数据.csv',encoding='utf-8')

#死循环接收消息
while True:
    #查找未读消息
    we = hw.TextControl(searchDepth=4)

    #死循环维持
    while not we.Exists(0):
        pass
    print('查找未读消息',we)
    #存在未读消息
    if we.Name:
        #点击未读消息
        we.Click(simulateMove=False)
        #读取最后一条消息
        last_msg=wx.ListControl(Name='消息').GetChildren()[-1].Name
        print('读取最后一条消息',last_msg)
        #判断关键字
        msg=df.apply(lambda x:x['回复内容'] if x['关键字'] in last_msg else None,axis=1)
        #数据筛选,移除空数据
        msg.dropna(axis=0,how='any',inplace=True)
        """
        dropna(): 这是 Pandas 中删除缺失数据的函数。
                axis=0: 这个参数指定操作的轴。axis=0 表示按行进行操作。
                如果是 axis=1，则表示按列进行操作。
                how='any': 这个参数决定删除的方式。
                how='any' 表示如果该行中任何一个元素是缺失值（NaN），就删除该行。
                如果是 how='all'，则表示只有当整行都是缺失值时才会删除该行。
                inplace=True: 这个参数表示是否在原地修改数据。
                如果是 True，则会直接修改 msg 对象；
                如果是 False，则返回一个新的 DataFrame，原始的 msg 不会被修改。
        
        """
        #做出列表
        ar=np.array(msg).tolist()
        #能够匹配到数据时
        if ar:
            #将数据输入
            #替换换行符号
            wx.SendKeys(ar[0].replace('{br}','{Shift}{Enter}'),waitTime=0)
            #发送消息
            wx.SendKeys('{Enter}',waitTime=0)
            #通过消息匹配检索会话栏的联系人
            wx.TextControl(SubName=ar[0][:5]).RightClick()
        else:
            wx.SendKeys('说啥呢',waitTime=0)
            wx.SendKeys('{Enter}',waitTime=0)
            wx.TextControl(SubName=last_msg[:5]).RightClick()

        #匹配右击控件
        ment = MenuControl(ClassName='CMenuWnd')
        #点击右键控件中的不显示聊天
        ment.TextControl(Name='不显示聊天').Click()