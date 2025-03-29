#-----------------------------------------------------------------------------------
# wxauto itself does not support adding friends in batches, so you need to modify the external library subfunctions
# After right clicking AddNewFriend, Go to Definition
# Add a line under self._show() in the AddNewFriend function: self.SwitchToChat() to add in batches
#-----------------------------------------------------------------------------------

import pandas as pd
from wxauto import WeChat
import time
import random


file_path = "List of people to be added.xlsx"  # Replace with your actual path
df = pd.read_excel(file_path, engine='openpyxl')

# Initialize WeChat
wx = WeChat()


# Traverse each row of data
for index, row in df.iterrows():
    try:
        # Extracting information
        phone = str(row['phone']).strip()
        addmsg = "Hello，xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"                       # The verification message sent when adding a friend has a character limit.
        remark = str(row['remark']).strip() if pd.notna(row['remark']) else None   # Note: You can set it in advance in the Excel list to be added
        raw_tags = str(row['tag']).strip() if pd.notna(row['tag']) else ""         # Friend tags can be set in advance in the Excel list to be added
        tags = [tag.strip() for tag in raw_tags.split(",") if tag.strip()]

        # Try adding friends
        start_time = time.time()
        success = wx.AddNewFriend(
            keywords=phone,
            addmsg=addmsg,
            remark=remark,
            tags=tags
        )

        # Determine the result based on the return value
        if success:
            print(f"✅ Successfully added:{phone}")
        else:
            print(f"❌ Add failed: {phone} (The phone number information is incorrect and cannot be added)")
        
        # Avoid frequent operations
        time.sleep(5 + random.uniform(1, 5))  

    except Exception as e:
        print(f"❌ Failed to add: {phone}, error: {str(e)}")
        continue
