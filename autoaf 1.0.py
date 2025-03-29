#-----------------------------------------------------------------------------------
# wxauto本身不支持批量添加好友，需要对外部库子函数进行修改
# 右键AddNewFriend后，Go to Definition
# 在AddNewFriend函数中的self._show()下添加一行：self.SwitchToChat()，就可以实现批量添加了
#-----------------------------------------------------------------------------------

import pandas as pd
from wxauto import WeChat
import time
import random


file_path = "待添加人员名单.xlsx"  # 替换为您的实际路径
df = pd.read_excel(file_path, engine='openpyxl')

# 初始化微信
wx = WeChat()


# 遍历每一行数据
for index, row in df.iterrows():
    try:
        # 提取信息
        phone = str(row['phone']).strip()
        addmsg = "您好，xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"                       # 添加好友时发送的验证消息，需注意有字数限制
        remark = str(row['remark']).strip() if pd.notna(row['remark']) else None  # 备注，可以在excel待添加名单中提前设置好
        raw_tags = str(row['tag']).strip() if pd.notna(row['tag']) else ""        # 好友标签，可以在excel待添加名单中提前设置好
        tags = [tag.strip() for tag in raw_tags.split(",") if tag.strip()]

        # 尝试添加好友
        start_time = time.time()
        success = wx.AddNewFriend(
            keywords=phone,
            addmsg=addmsg,
            remark=remark,
            tags=tags
        )

        # 根据返回值判断结果
        if success:
            print(f"✅ 成功添加：{phone}")
        else:
            print(f"❌ 添加失败：{phone}（电话号信息有误，无法添加）")
        
        # 避免频繁操作
        time.sleep(5 + random.uniform(1, 5))  

    except Exception as e:
        print(f"❌ 添加失败：{phone}，错误：{str(e)}")
        continue