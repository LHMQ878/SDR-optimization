import requests
import json
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime, timezone

# Configuration parameters
CONFIG = {
    "API_KEY": "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",        # API key for Silicon Flow Large Model
    "MODEL_NAME": "deepseek-ai/DeepSeek-V3",                          
    "API_URL": "https://api.siliconflow.cn/v1/chat/completions",
    "APP_ID": "cli_xxxxxxxxxxxxxxxx",                                 # Feishu Enterprise Self-built Robot Verification Information
    "APP_SECRET": "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",                 # Same as above
    "APP_TOKEN": "xxxxxxxxxxxxxxxxxxxxxxxxxxx",                       # The unique identifier of the multidimensional table app
    "TABLE_ID_AIGC": "xxxxxxxxxxxxxxxx",                              # The unique identifier of the multidimensional table data table.
    "TABLE_ID_SC": "xxxxxxxxxxxxxxxx",
    "TABLE_ID_GS": "xxxxxxxxxxxxxxxx"
}

class FeishuProcessor:
    def __init__(self):
        self.access_token = get_feishu_token(CONFIG["APP_ID"], CONFIG["APP_SECRET"])
        self.current_course_type = None
        self.window = None  # 用于存储当前窗口的引用
        self.processing_label = None  # 用于存储处理中的提示标签

    def create_type_selection_window(self):
        root = tk.Tk()
        root.title("选择课程类型")
        root.geometry("380x250")

        tk.Label(root, text="请选择要处理的会议记录类型").pack(pady=20)
        tk.Button(root, text="AI通识课", command=lambda: self.create_input_window("AIGC", root)).pack(pady=10)
        tk.Button(root, text="专业课", command=lambda: self.create_input_window("SC", root)).pack(pady=10)
        tk.Button(root, text="泛科研情况", command=lambda: self.create_input_window("GS", root)).pack(pady=10)

        root.mainloop()

    def create_input_window(self, course_type, parent_window):
        parent_window.destroy()
        self.current_course_type = course_type
        
        self.window = tk.Tk()  # 将窗口引用保存到实例变量中
        self.window.title(f"{course_type} 会议纪要处理")
        self.window.geometry("600x400")

        # 界面组件
        tk.Label(self.window, text="请选择输入方式：").pack(pady=10)
        self.upload_button = tk.Button(self.window, text="上传txt文件", command=self.upload_file)
        self.upload_button.pack(pady=5)
        
        tk.Label(self.window, text="或直接输入会议纪要内容：").pack(pady=10)
        self.text_area = tk.Text(self.window, wrap=tk.WORD, height=10)
        self.text_area.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        
        self.submit_button = tk.Button(self.window, text="提交处理", command=self.process_text_input)
        self.submit_button.pack(pady=10)
        
        self.window.mainloop()

    def upload_file(self):
        # 禁用上传按钮，防止重复上传
        self.upload_button.config(state=tk.DISABLED)
        
        # 显示文件上传成功的提示
        if self.processing_label:
            self.processing_label.destroy()  # 移除之前的提示
        self.processing_label = tk.Label(self.window, text="文件已经上传成功，正在处理中...", fg="blue")
        self.processing_label.pack(pady=10)
        
        # 强制更新界面
        self.window.update()
        
        # 选择文件并读取内容
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                self.process_content(content)
            except Exception as e:
                messagebox.showerror("错误", f"读取文件失败: {str(e)}")
                self.reset_ui_state()  # 重置UI状态
        else:
            self.reset_ui_state()  # 如果未选择文件，重置UI状态

    def process_text_input(self):
        content = self.text_area.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showwarning("警告", "请输入会议纪要内容")
            return
        
        # 禁用提交按钮，防止重复提交
        self.submit_button.config(state=tk.DISABLED)
        
        # 显示处理中的提示
        if self.processing_label:
            self.processing_label.destroy()  # 移除之前的提示
        self.processing_label = tk.Label(self.window, text="正在处理中，请稍等...", fg="blue")
        self.processing_label.pack(pady=10)
        
        # 强制更新界面
        self.window.update()
        
        # 开始处理内容
        self.process_content(content)

    def process_content(self, content):
        try:
            if self.current_course_type == "AIGC":
                result = call_deepseek_v3_AIGC(
                    content, 
                    CONFIG["API_KEY"], 
                    CONFIG["MODEL_NAME"], 
                    CONFIG["API_URL"]
                )
                table_id = CONFIG["TABLE_ID_AIGC"]
            elif self.current_course_type == "SC":
                result = call_deepseek_v3_SC(
                    content, 
                    CONFIG["API_KEY"], 
                    CONFIG["MODEL_NAME"], 
                    CONFIG["API_URL"]
                )
                table_id = CONFIG["TABLE_ID_SC"]
            elif self.current_course_type == "GS":
                result = call_deepseek_v3_GS(
                    content,
                    CONFIG["API_KEY"],
                    CONFIG["MODEL_NAME"],
                    CONFIG["API_URL"]
                )
                table_id = CONFIG["TABLE_ID_GS"]
            else:
                messagebox.showerror("错误", "未选择课程类型")
                return

            if result:
                processed_data = process_json_data(result)
                if processed_data:
                    success = write_to_feishu_table(
                        processed_data,
                        CONFIG["APP_TOKEN"],
                        table_id,
                        self.access_token
                    )
                    if success:
                        messagebox.showinfo("成功", "数据已成功写入飞书表格！")
                    else:
                        messagebox.showerror("错误", "写入飞书表格失败")
            else:
                messagebox.showerror("错误", "模型调用失败，请检查API或网络连接")
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中发生错误: {str(e)}")
        finally:
            self.reset_ui_state()  # 重置UI状态

    def reset_ui_state(self):
        """重置UI状态，恢复按钮并移除提示"""
        if self.processing_label:
            self.processing_label.destroy()
            self.processing_label = None
        self.submit_button.config(state=tk.NORMAL)
        self.upload_button.config(state=tk.NORMAL)
        self.window.update()

# 获取飞书访问令牌
def get_feishu_token(app_id, app_secret):
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json; charset=utf-8"}
    payload = {"app_id": app_id, "app_secret": app_secret}
    response = requests.post(url, headers=headers, json=payload)
    return response.json().get("tenant_access_token")

# 调用DeepSeek-V3模型生成AI通识课的结构化JSON信息
def call_deepseek_v3_AIGC(prompt, api_key, model_name, api_url, retries=5, delay=10):
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": "你是一个专业的会议纪要助手，能够从会议记录中提取结构化数据"},
            {"role": "user", "content": f"""
                请从以下会议记录中提取结构化数据，按JSON格式返回：
                - 会议主题（key: 会议主题, 具体某学校&和鲸交流会）
                - 时间（key: 时间，格式为YYYY-MM-DD）
                - 学科类别（key: 学科类别, 选项：综合, 理科, 工科, 医学, 经济管理, 人文社科, 农学, 文学）
                - 课程预期效果（key: 课程预期效果, 客户原话）
                - 牵头部门（key: 牵头部门, 计算机学院或教务处...）
                - 方案设计部门（key: 方案设计部门）
                - 采购决策部门（key: 采购决策部门）
                - 课程规模（key: 课程规模, 进一步补充并发人数）
                - 学生专业（key: 学生专业）
                - 学生年级（key: 学生年级）
                - 开课时间（key: 开课时间，选项: 2025年春季学期, 2025年秋季学期, 2026年春季学期, 2026年秋季学期, 其他）
                - 开课形式（key: 开课形式）
                - 人才培养方案（key: 人才培养方案, 主要了解通识课是否规划进了人才培养）
                - 期望优化要点（key: 期望优化要点，如果已经开课需要重点了解期望优化的要点，是否与我们的优势相匹配）
                - 课程名称（key: 课程名称）
                - 客户原本是否有课件（key: 客户原本是否有课件, 选项: "是", "否，需要用我们的课件", "自己有课件也需要我们的课件"）
                - 现有课件来源（key: 现有课件来源，判断是否需要我们的课程）
                - 教材配套情况和需求（key: 教材配套情况和需求）
                - 理论教学方式（Key: 理论教学方式）
                - 实践教学方式（Key: 实践教学方式, 期望的实践形式能否契合我们的优势）
                - 实践教学占比（key: 实践教学占比）
                - 实践教学难度与形式期望（key: 实践教学难度与形式期望）
                - 实验案例资源情况（key: 实验案例资源情况, 有没有, 期望有哪些）
                - 部署方式（key: 部署方式）
                - 服务器情况（key: 服务器情况, 算力资源）
                - 是否有相关平台（key: 是否有相关平台，已有的话判断竞品优劣势、对modelwhale的需求程度）
                - 场地机房情况（key: 场地机房情况，是希望集中起来上课还是能够随时随地上课）
                - 决策链（key: 决策链, 是否有较为明确的经费来源/预算）

                会议记录：
                {prompt}
            """}
        ],
        "temperature": 0.3,
        "max_tokens": 500,
        "stream": False
    }
    for i in range(retries):
        try:
            response = requests.post(api_url, headers=headers, json=payload, timeout=100)
            response.raise_for_status()
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            content = content.replace("```json", "").replace("```", "").strip()
            if content.strip().startswith(('{', '[')):
                return content
            else:
                print(f"返回的内容可能不是有效的JSON格式: {content}")
                return None
        except requests.exceptions.RequestException as e:
            print(f"第 {i + 1} 次尝试失败: {e}")
            time.sleep(delay)
        except (KeyError, IndexError):
            print(f"响应数据格式不符合预期: {response.text}")
            return None
    return None

# 调用DeepSeek-V3模型生成专业课的结构化JSON信息
def call_deepseek_v3_SC(prompt, api_key, model_name, api_url, retries=5, delay=10):
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": "你是一个专业的会议纪要助手，能够从会议记录中提取结构化数据。"},
            {"role": "user", "content": f"""
                请从以下会议记录中提取结构化数据，按JSON格式返回：
                - 会议主题（key: 会议主题, 具体某学校&和鲸交流会）
                - 时间（key: 时间，格式为YYYY-MM-DD）
                - 学科类别（key: 学科类别, 选项：地球科学, 气象, 计算机, 医学, 经济管理, 人文社科, 农学, 其他）
                - 学校重点信息（key: 学校重点信息, 来源于输入的背景调查）
                - 学科评估级别（key: 学科评估级别, 选项：A+, A, A-, B+, B, B-, C+, C, C-)
                - 学校教改信息（key: 学校教改信息, 来源于输入的背景调查）
                - 是否发过相关论文或相关课题（key: 是否发过相关论文或相关课题, 来源于输入的背景调查）
                - 老师身份（key: 老师身份, 来源于输入的背景调查或者与老师沟通获取）
                - 老师编程能力（key: 老师编程能力, 来源于输入的背景调查或者与老师沟通获取）
                - 具体专业名称（key: 具体专业名称）
                - 专业开设时间（key: 专业开设时间）
                - 课程情况（key: 课程情况, 需要具体名称，是否有具体的数据分析、人工智能结合的课程）
                - 课程开设时间（key: 课程开设时间, 开了多久，或者预计开课时间）
                - 是否有完整课件数据代码（key: 是否有完整课件数据代码）
                - 已有平台情况（key: 已有平台情况）
                - 新增平台的预期使用计划（key: 新增平台的预期使用计划, 具体什么人、在什么场景下使用，想达到什么效果，例如给大三50人XX课上实验课）
                - 期待部署方式（key: 期待部署方式, 选项: 本地化部署, 公有云）
                - 学院是否已经或有计划买服务器和相关资源（key: 学院是否已经或有计划买服务器和相关资源）
                - 潜在实验室建设需求摸排与记录（key: 潜在实验室建设需求摸排与记录）
                - 决策链（key: 决策链, 是否有较为明确的经费来源/预算）

                会议记录：
                {prompt}
            """}
        ],
        "temperature": 0.3,
        "max_tokens": 500,
        "stream": False
    }
    for i in range(retries):
        try:
            response = requests.post(api_url, headers=headers, json=payload, timeout=100)
            response.raise_for_status()
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            content = content.replace("```json", "").replace("```", "").strip()
            if content.strip().startswith(('{', '[')):
                return content
            else:
                print(f"返回的内容可能不是有效的JSON格式: {content}")
                return None
        except requests.exceptions.RequestException as e:
            print(f"第 {i + 1} 次尝试失败: {e}")
            time.sleep(delay)
        except (KeyError, IndexError):
            print(f"响应数据格式不符合预期: {response.text}")
            return None
    return None

# 调用DeepSeek-V3模型生成泛科研情况的结构化JSON信息
def call_deepseek_v3_GS(prompt, api_key, model_name, api_url, retries=5, delay=10):
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": "你是一个专业的会议纪要助手，能够从会议记录中提取结构化数据。"},
            {"role": "user", "content": f"""
                请从以下会议记录中提取结构化数据，按JSON格式返回：
                - 客户单位（key: 客户单位）
                - 时间（key: 时间，格式为YYYY-MM-DD）
                - 学科专业领域（key: 学科专业领域, 具体研究方向）
                - 线索来源说明（key: 线索来源说明, 非一级学科、需为具体研究方向）
                - 对接人姓名（key: 对接人姓名）
                - 所在部门与职务/负责事项（key: 所在部门与职务/负责事项）
                - 技术能力（key: 技术能力）
                - 所使用的科研分析工具（key: 所使用的科研分析工具, "多选":["jupyterlab", "pycharm", "vscode", "matlab", "spss", "Rstudio", "其他专业软件"]）
                - 产品试用情况（key: 产品试用情况, 选项: "已运行", "已开通-未使用", "已运行"）
                - 经费来源（key: 经费来源）
                - 需求产品（key: 需求产品，"多选":["教学实训", "科学分析", "机器学习", "大模型应用", "生态服务"]）
                - 数据保密情况（key: 数据保密情况, 选项: "是", "否"）
                - 数据来源（key: 数据来源, "多选":["单位内部", "网络公开数据", "导师、同事等提供", "个人收集", "其他"]）
                - 涉及到的数据类型（key: 涉及到的数据类型）
                - 数据分享使用工具（key: 数据分享使用工具, "多选":["微信、飞书等通讯工具", "网盘", "共享服务器", "其他"]）
                - 算力情况（key: 算力情况）
                - 日常算力使用申请（key: 日常算力使用申请，选项: "需要申请", "不需要申请", "资源紧张", "资源充足", "仅使用自己电脑"）
                - AI算法模型团队人数（Key: AI算法模型团队人数）
                - 模型算法成果积累（Key: 模型算法成果积累, "多选":["自己开发的", "团队开发的", "外部合作开发的"]）
                - 模型算法成果具体应用说明（key: 模型算法成果具体应用说明,  目前业务中经常用到的、举例说明具体场景）
                - 计划使用哪些大模型（key: 计划使用哪些大模型）
                - 大模型规划情况（key: 大模型规划情况, 单位的态度、对接人的态度）
                - 是否使用大模型（key: 是否使用大模型, 具体是哪些，使用的场景是？）
                - 大模型人才团队建设情况（key: 大模型人才团队建设情况）

                会议记录：
                {prompt}
            """}
        ],
        "temperature": 0.3,
        "max_tokens": 500,
        "stream": False
    }
    for i in range(retries):
        try:
            response = requests.post(api_url, headers=headers, json=payload, timeout=100)
            response.raise_for_status()
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            content = content.replace("```json", "").replace("```", "").strip()
            if content.strip().startswith(('{', '[')):
                return content
            else:
                print(f"返回的内容可能不是有效的JSON格式: {content}")
                return None
        except requests.exceptions.RequestException as e:
            print(f"第 {i + 1} 次尝试失败: {e}")
            time.sleep(delay)
        except (KeyError, IndexError):
            print(f"响应数据格式不符合预期: {response.text}")
            return None
    return None

# 处理模型返回的JSON数据
def process_json_data(json_str):
    try:
        json_data = json.loads(json_str)
        if "时间" in json_data:
            try:
                date_obj = datetime.strptime(json_data["时间"], "%Y-%m-%d")
                utc_date = date_obj.replace(tzinfo=timezone.utc)
                json_data["时间"] = int(utc_date.timestamp() * 1000)
            except ValueError:
                print("时间格式转换失败，请检查时间格式。")
                json_data["时间"] = None
        return json_data
    except json.JSONDecodeError as e:
        print(f"结果解析失败: {e}，返回的内容可能不是有效的JSON格式。")
        return None

# 写入数据到飞书表格
def write_to_feishu_table(data, app_token, table_id, access_token):
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    payload = {"fields": data}
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=10)
        response.raise_for_status()
        response_json = response.json()
        if "data" in response_json and "record" in response_json["data"]:
            print("写入成功！记录ID:", response_json["data"]["record"]["record_id"])
            return True
        else:
            print("响应格式异常，缺少关键字段:", response_json)
            return False
    except requests.exceptions.RequestException as e:
        print(f"请求失败: {str(e)}")
        if response:
            print(f"错误响应内容: {response.text}")
    except json.JSONDecodeError:
        print("响应解析失败，返回内容不是有效JSON")
    except Exception as e:
        print(f"未知错误: {str(e)}")
    return False


if __name__ == "__main__":
    processor = FeishuProcessor()
    processor.create_type_selection_window()