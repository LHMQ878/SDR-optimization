# 标准库
import os
import sys
import time
import threading
import logging
from queue import Queue
from datetime import datetime
from typing import Tuple, Optional, List

# 第三方库
import pyautogui
import requests
import tiktoken
from docx import Document
from rich.console import Console
from rich.layout import Layout
from rich.live import Live
from rich.panel import Panel
from rich.progress import Progress
from rich.text import Text

# 本地库
from wxauto import WeChat

# 配置日志系统
logging.basicConfig(
    filename='wechat_assistant.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

# ====================
# 配置区域（用户可修改）
# ====================
class Config:
    # DeepSeek API 配置 
    API_URL = "https://api.siliconflow.cn/v1/chat/completions"
    API_KEYS = [
        "sk-fiuiulyyxuutbvxokyqpboiitbjpohpfxxmwrfbhyxxxxxxx",  # 主KEY
        "sk-drsxypzgdhptprgtkprdyfihdcvmpxwcxalwdhmvdxxxxxxx",  # 备用KEY1
        "sk-gnsgxoazmujgptmlrdttovajtsotpuiemivelvenrxxxxxxx",  # 备用KEY2
        "sk-qicoopotuposhfudroszcqybxxepjlkmlpqiezjoxxxxxxxx",  # 备用KEY3
        "sk-ohlhuntbdwdocylzzzxifapjmehrboxwksljlkgaixxxxxxx"   # 备用KEY4
    ]

    # 监听配置
    LISTEN_NAMES = ['群聊名称']                     # 监听的微信联系人/群名称
    CHECK_INTERVAL = 1                            # 新消息检查间隔（秒）
    MAX_DURATION = 300                            # 最大运行时间（秒）

    # 知识库配置（改为文件夹形式）
    KNOWLEDGE_DIR = r"D:\heywhale\WeChat Agent\KNOWLEDGE_SOURCE"  # 使用原始字符串处理路径
    SUPPORTED_EXTS = [".docx", ".txt", ".pdf"]  # 支持docx, txt, pdf
    MAX_TOKENS = 25600 

    # 重试配置
    API_MAX_RETRIES = 5                     # 最大重试次数
    API_RETRY_DELAY = 1                     # 初始延迟（秒）
    API_RETRY_STATUS_CODES = [429, 502]     # 需要重试的状态码
    API_KEY_ROTATE_CODES = [403, 429]        # 需要切换KEY的状态码
    _current_key_index = 0                  # 当前使用的KEY索引
    _key_lock = threading.Lock()            # KEY切换锁
 

# ====================
# 初始化全局组件
# ====================
console = Console()
message_queue = Queue()  # 用于线程间通信
wx = WeChat()            # 微信客户端

class APIClient:
    @staticmethod
    def get_auth_header() -> dict:
        """获取当前API密钥的认证头"""
        with Config._key_lock:
            key = Config.API_KEYS[Config._current_key_index % len(Config.API_KEYS)]
        return {"Authorization": f"Bearer {key}"}

    @staticmethod
    def rotate_key():
        """轮换到下一个API密钥"""
        with Config._key_lock:
            Config._current_key_index += 1
            logging.warning(f"切换到备用KEY #{Config._current_key_index % len(Config.API_KEYS) + 1}")

class KnowledgeManager:
    """知识库管理器：整合多来源知识"""
    
    @staticmethod
    def calculate_tokens(text: str) -> int:
        """计算文本的Token数量"""
        try:
            # 尝试使用DeepSeek专用编码（如果不存在则使用默认）
            encoder = tiktoken.encoding_for_model("DeepSeek-R1-Distill-Qwen-32B")
        except KeyError:
            encoder = tiktoken.get_encoding("cl100k_base")
        return len(encoder.encode(text))

    @staticmethod
    def load_folder() -> Tuple[str, int]:
        """加载知识库并返回（内容，总Token数）"""
        combined: List[str] = []
        total_tokens = 0
        
        try:
            if not os.path.exists(Config.KNOWLEDGE_DIR):
                console.print(f"[red]❌ 知识库文件夹不存在: {Config.KNOWLEDGE_DIR}[/]")
                return "", 0

            for root, _, files in os.walk(Config.KNOWLEDGE_DIR):
                for file in files:
                    file_ext = os.path.splitext(file)[1].lower()
                    if file_ext in Config.SUPPORTED_EXTS:
                        file_path = os.path.join(root, file)
                        content = KnowledgeManager.load(file_path)
                        
                        if content:
                            # 计算并累加Token
                            tokens = KnowledgeManager.calculate_tokens(content)
                            total_tokens += tokens
                            console.print(f"[dim]📄 {os.path.basename(file_path)}: {tokens} tokens[/]")
                            
                            # 实时检查Token限制
                            if total_tokens > Config.MAX_TOKENS:
                                console.print(f"[red]❌ 知识库Token超过限制 ({total_tokens}/{Config.MAX_TOKENS})[/]")
                                sys.exit(1)
                                
                            # 错误代码（KnowledgeManager.load_folder）
                            combined.append(f"【知识来源：os.path.basename(file_path)】\n{content}")

            console.print(f"[green]✅ 知识库加载完成 总Token数: {total_tokens}/{Config.MAX_TOKENS}[/]")
            return "\n\n".join(combined), total_tokens
            
        except Exception as e:
            console.print(f"[red]❌ 加载知识库失败: {str(e)}[/]")
            return "", 0

    @staticmethod
    def count_files() -> int:
        """统计知识库文件夹中支持的文件数量"""
        count = 0
        try:
            if not os.path.exists(Config.KNOWLEDGE_DIR):
                return 0
                
            for root, _, files in os.walk(Config.KNOWLEDGE_DIR):
                for file in files:
                    file_ext = os.path.splitext(file)[1].lower()
                    if file_ext in Config.SUPPORTED_EXTS:
                        count += 1
            return count
        except Exception as e:
            console.print(f"[red]❌ 统计文件失败: {str(e)}[/]")
            return 0

    @staticmethod
    def load(file_path: str) -> str:
        """根据文件类型加载单个文件内容"""
        try:
            if file_path.endswith(".docx"):
                return KnowledgeManager._load_docx(file_path)
            elif file_path.endswith(".txt"):
                return KnowledgeManager._load_txt(file_path)  # 新增txt处理
            elif file_path.endswith(".pdf"):
                return KnowledgeManager._load_pdf(file_path)
            else:
                console.print(f"[yellow]⚠️ 跳过不支持的文件类型: {file_path}[/]")
                return ""
        except Exception as e:
            console.print(f"[red]❌ 加载失败 {file_path}: {str(e)}[/]")
            return ""
    
    @staticmethod
    def _load_docx(path: str) -> str:
        """读取Word文档内容"""
        try:
            doc = Document(path)
            return "\n".join([p.text for p in doc.paragraphs])
        except Exception as e:
            console.print(f"[red]❌ 读取Word失败 {path}: {str(e)}[/]")
            return ""

    @staticmethod
    def _load_txt(path: str) -> str:  
        """读取文本文件内容"""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            console.print(f"[red]❌ 读取TXT失败 {path}: {str(e)}[/]")
            return ""
        
    @staticmethod
    def _load_pdf(path: str) -> str:
        """读取PDF文件内容"""
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            console.print("[red]❌ 请先安装PyPDF2库：pip install PyPDF2[/]")
            return ""

        text = ""
        try:
            with open(path, 'rb') as file:
                pdf_reader = PdfReader(file)
                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
            return text.strip()
        except Exception as e:
            console.print(f"[red]❌ 读取PDF失败 {path}: {str(e)}[/]")
            return ""

class ChatLogger:
    """聊天日志记录器"""
    
    def __init__(self):
        self.log = []
        
    def add_entry(self, sender: str, message: str, reply: str, error: Optional[str] = None):
        """添加日志条目"""
        entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "sender": sender,
            "message": message,
            "reply": reply,
            "error": error
        }
        self.log.append(entry)
        # console.print(f"[dim]📝 已记录：{entry}[/]")
        
    def save_to_file(self):
        """保存日志到TXT文件，文件名包含结束时间"""
        try:
            # 获取当前脚本文件的目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            
            # 动态生成日志文件名
            end_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file_name = f"chat_log_{end_time}.txt"
            
            # 构建完整路径
            log_file_path = os.path.join(current_dir, log_file_name)
            
            # 写入日志文件
            with open(log_file_path, "w", encoding="utf-8") as f:
                f.write("微信智能助手对话日志\n")
                f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 50 + "\n\n")
                
                for i, entry in enumerate(self.log, 1):
                    f.write(f"对话记录 #{i}\n")
                    f.write(f"时间: {entry['timestamp']}\n")
                    f.write(f"发送者: {entry['sender']}\n")
                    f.write(f"原始消息: {entry['message']}\n")
                    f.write(f"回复内容: {entry['reply']}\n")
                    f.write("-" * 50 + "\n")
                    
            console.print(f"[green]✅ 日志已保存至 {log_file_path}[/]")
        except Exception as e:
            console.print(f"[red]❌ 日志保存失败: {str(e)}[/]")

# ====================
# 核心功能模块
# ====================
class WeChatAssistant:
    """微信智能助手主程序"""
    
    def __init__(self):
        # 加载知识库并检查Token
        self.knowledge, self.knowledge_tokens = KnowledgeManager.load_folder()
        
        if self.knowledge_tokens == 0:
            console.print("[yellow]⚠️ 知识库为空，将仅使用基础模型[/]")
        elif self.knowledge_tokens > Config.MAX_TOKENS:
            console.print(f"[red]❌ 知识库Token数超过限制 ({self.knowledge_tokens}/{Config.MAX_TOKENS})[/]")
            sys.exit(1)  
        self.logger = ChatLogger()
        self.last_received = "暂无消息"
        self.last_reply = "暂无回复"
        self.lock = threading.Lock()
        
        
    def _load_knowledge(self) -> str:
        """整合所有知识源"""
        combined = []
        for source in Config.KNOWLEDGE_SOURCES:
            content = KnowledgeManager.load(source)
            if content:
                combined.append(f"【知识来源：{os.path.basename(source)}】\n{content}")
        return "\n\n".join(combined)
    
    def _call_ai_api(self, prompt: str) -> str:
        """调用大语言模型API"""
        # console.print(f"[blue]🤖 生成回复中 | 输入Token: {KnowledgeManager.calculate_tokens(prompt)}[/]")

        system_prompt = f"""请根据以下知识库回答问题（如果问题超出知识库范围，可以结合常识进行推理回答）：
        {self.knowledge}
        回答要求：
        1. 语气自然：用口语化的中文回答，像朋友聊天一样轻松自然，避免过于正式或机械的表达。
        2. 适当幽默：如果问题适合，可以加入一些幽默或轻松的语气，让对话更有趣。
        3. 表情和语气词：如果对方使用了表情或语气词，可以适当回应类似的表情或语气词，增加互动感。
        4. 委婉拒绝：如果问题超出知识库范围，可以委婉地表示“这个问题我不太清楚哦”或“我可能需要再学习一下”，避免直接说“无法回答”。
        5. 简洁明了：回答要言简意赅，避免冗长的解释，也不要输出推理过程，保持对话流畅。
        6. 情感共鸣：如果问题涉及情感或主观感受，可以适当表达理解或共情，比如“我明白你的感受”或“这确实是个有趣的想法”。
        """
        
        '''headers = {
            "Authorization": f"Bearer {Config.API_KEY}",
            "Content-Type": "application/json"
        }'''
        
        data = {
            "model": "deepseek-ai/DeepSeek-V3",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.3,
            "max_tokens": 300
        }
        
        # 指数退避重试逻辑
        retries = 0
        delay = Config.API_RETRY_DELAY
        last_status = None

        try:
            while retries < Config.API_MAX_RETRIES:
                headers = APIClient.get_auth_header()
                headers["Content-Type"] = "application/json"

                try:
                    response = requests.post(
                        Config.API_URL,
                        headers=headers,
                        json=data,
                        timeout=45
                    )
                    response.raise_for_status()
                    return response.json()["choices"][0]["message"]["content"]

                except requests.exceptions.HTTPError as e:
                    status_code = e.response.status_code
                    last_status = status_code
                    
                    # 需要切换KEY的错误
                    if status_code in Config.API_KEY_ROTATE_CODES:
                        APIClient.rotate_key()
                    
                    # 需要重试的错误
                    if status_code in Config.API_RETRY_STATUS_CODES:
                        logging.warning("API错误 %d, %d/%d次重试...", status_code, retries+1, Config.API_MAX_RETRIES)
                        time.sleep(delay)
                        delay *= 2  # 指数退避
                        retries += 1
                    else:
                        break

                except (requests.exceptions.ConnectionError, 
                        requests.exceptions.Timeout) as e:
                    logging.warning("网络错误: %s, %d/%d次重试...", 
                                 str(e), retries+1, Config.API_MAX_RETRIES)
                    last_status = "NETWORK_ERROR"  # 标记网络错误
                    time.sleep(delay)
                    delay *= 2
                    retries += 1

            # 重试结束后使用 last_status
            if last_status is not None:
                logging.error("所有重试失败，最终状态码: %s", last_status)
            else:
                logging.error("未知错误导致请求失败")

            return "服务暂时不可用，请稍后再试"  # 兜底回复
                    
        except Exception as e:
            logging.exception("API调用失败")
            return None

    def _handle_message(self, chat, msg):
        """添加详细日志和异常处理"""
        try:
            console.print(f"[red]🔥 收到消息调试标记[/]")
            sender = msg.sender
            message = msg.content
            # console.print(f"\n[cyan]📩 新消息 @{datetime.now().strftime('%H:%M:%S')}[/]")
            # console.print(f"发件人: {sender}\n内容: {message}\n")
            
            reply = self._call_ai_api(message)
            error = None if reply else "API调用失败"

            if error:
                console.print(f"[red]⚠️ 生成回复失败: {error}[/]")
                reply = "抱歉，暂时无法处理您的请求。"
                
            #console.print(f"[green]💬 生成回复:\n{reply}[/]")
            
            # 发送消息
            with self.lock:
                chat.SendMsg(reply)
                # console.print(f"[green]📨 已发送回复 @{datetime.now().strftime('%H:%M:%S')}[/]")
                self.logger.add_entry(sender, message, reply, error=error)
                message_queue.put((f"[{sender}] {message}", reply))
            
        except Exception as e:
            logging.exception("消息处理异常")
            import traceback
            traceback.print_exc()

    def _setup_ui(self) -> Layout:
        """初始化终端界面布局"""
        layout = Layout()
        layout.split(
            Layout(name="header", size=3),
            Layout(name="main", ratio=1),
            Layout(name="footer", size=5)
        )
        layout["main"].split_row(
            Layout(name="messages", ratio=2),
            Layout(name="status")
        )
        return layout

    def _update_ui(self, layout: Layout, start_time: float):
        """刷新终端界面显示"""
        # Header
        elapsed = time.time() - start_time
        remaining = Config.MAX_DURATION - elapsed
        header_text = Text(f" 基于RAG的智能问答助手 | 运行: {time.strftime('%H:%M:%S', time.gmtime(elapsed))} | 剩余: {remaining:.0f}s", 
                        style="bold white on blue")
        layout["header"].update(Panel(header_text))

        # Messages
        msg_content = Text()
        msg_content.append(f"最后收到的消息:\n{self.last_received}\n\n", style="cyan")
        msg_content.append(f"最后发送的回复:\n{self.last_reply}", style="green")
        layout["messages"].update(Panel(msg_content, title="消息记录"))

        # Status
        status_content = Text()
        status_content.append(f"当前时间: {datetime.now().strftime('%H:%M:%S')}\n", style="yellow")
        status_content.append(f"监听窗口: {', '.join(Config.LISTEN_NAMES)}\n")
        file_count = KnowledgeManager.count_files()
        status_content.append(f"知识库来源: {file_count}个文件\n")
        status_content.append(f"知识库Token: {self.knowledge_tokens}/{Config.MAX_TOKENS}\n")
        layout["status"].update(Panel(status_content, title="系统状态"))

        # Footer
        progress = Progress()
        progress.add_task("[cyan]运行进度", total=Config.MAX_DURATION, completed=elapsed)
        layout["footer"].update(Panel(progress, title="运行进度")) 

    def run(self):
        """启动主程序"""
        # 初始化监听
        for name in Config.LISTEN_NAMES:
            try:
                wx.ChatWith(who=name)
                current_x, current_y = pyautogui.position()
                pyautogui.moveTo(current_x, current_y - 30)
                pyautogui.doubleClick()
                wx.AddListenChat(who=name)
                console.print(f"[green]✅ 已锁定窗口: {name}[/]")
            except Exception as e:
                console.print(f"[red]❌ 窗口初始化失败: {name} - {str(e)}[/]")

        # 准备UI
        ui_layout = self._setup_ui()
        start_time = time.time()

        with Live(ui_layout, refresh_per_second=2, screen=True):
            while (time.time() - start_time) < Config.MAX_DURATION:
                try:
                    # 更新最后的消息记录
                    if not message_queue.empty():
                        self.last_received, self.last_reply = message_queue.get()

                    # 刷新界面
                    self._update_ui(ui_layout, start_time)

                    # 检查新消息
                    msgs = wx.GetListenMessage()
                    for chat in msgs:
                        for msg in msgs[chat]:
                            if msg.type == 'friend':
                                thread = threading.Thread(
                                    target=self._handle_message,
                                    args=(chat, msg)
                                )
                                thread.start()

                    time.sleep(Config.CHECK_INTERVAL)
                
                except KeyboardInterrupt:
                    console.print("[yellow]⏹️ 用户手动终止[/]")
                    break
                except Exception as e:
                    console.print(f"[red]⚠️ 异常: {str(e)}[/]")
                    time.sleep(1)

        # 保存日志
        self.logger.save_to_file()
        console.print(f"[green]⏰ 服务已安全停止，累计运行 {Config.MAX_DURATION}秒[/]")

# ====================
# 程序启动
# ====================
if __name__ == "__main__":
    try:
        # 检查tiktoken依赖
        import tiktoken
    except ImportError:
        console.print("[red]❌ 需要安装tiktoken库：pip install tiktoken[/]")
        sys.exit(1)

    assistant = WeChatAssistant()
    assistant.run()
