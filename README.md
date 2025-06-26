# Financial Analyst Agent - 智能金融分析师


**智能金融分析师**是一个基于大型语言模型（LLM）的AI Agent系统，旨在全流程自动化生成深度金融研究报告。它能够模拟人类专家的研究范式，自主完成从**实时信息检索**、**多源数据处理**、**深度分析**到**图文并茂报告生成**的完整任务闭环。


## 核心功能

- **自主决策工作流**: Agent能够基于预设的研究框架，动态决策下一步行动：是继续深入搜索，还是开始撰写章节，或是完成报告，实现了真正的研究流程自动化。
- **实时网络信息检索**: 集成 `DuckDuckGo` 搜索引擎，使Agent能够获取最新的宏观经济数据、政策动态和市场新闻，克服了大型语言模型知识时效性的限制。
- **多模态报告生成**: Agent不仅能撰写专业的分析文本，还能调用 `matplotlib` 将核心数据自动渲染成图表，并最终将所有内容整合输出为图文并茂的 `.docx` 报告。
- **高度模块化与可扩展**: 系统基于`PocketFlow`思想构建，将报告生成流程拆分为独立的`Node`（如决策、搜索、总结、生成），易于维护和功能扩展。


## 技术栈

- **核心框架**: Python 3.9+
- **LLM API**: OpenAI
- **工作流**: Custom PocketFlow-style Framework
- **数据处理与可视化**: Pandas, Matplotlib
- **网络检索**: duckduckgo-search
- **文档生成**: python-docx
- **环境管理**: python-dotenv

## 安装与配置

1.  **克隆仓库**
    ```bash
    git clone <your-repo-url>
    cd <your-repo-directory>
    ```

2.  **创建虚拟环境并激活**
    ```bash
    python -m venv venv
    # Windows
    .\venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```

3.  **安装依赖**
    ```bash
    pip install -r requirements.txt
    ```

4.  **配置环境变量**

    在项目根目录下创建一个名为 `.env` 的文件，并填入您的API密钥等信息。

    ```dotenv
    # .env file
    OPENAI_API_KEY="sk-..."
    OPENAI_BASE_URL="https://api.openai.com/v1" # 如果使用代理或第三方服务，请修改此项
    OPENAI_MODEL="gpt-4-turbo" # 您希望使用的模型
    ```


## 项目结构

```
.
├── data/              # 存放原始数据或知识库文件
├── reports/           # 存放最终生成的报告
├── report_dataframes/ # 存放报告生成过程中的数据表格
├── src/               # 项目核心源码
│   ├── agents/        # Agent角色定义
│   ├── tools/         # Agent可使用的工具 (如文档处理)
│   ├── workflows/     # 核心工作流定义 (如宏观分析流程)
│   └── frameworks/    # 自定义框架代码 (如PocketFlow)
├── .env               # 环境变量配置文件 (需自行创建)
├── create_reference.py# 创建知识库参考文件的脚本
├── main.py            # 主程序入口
├── requirements.txt   # 项目依赖
└── README.md          # 项目说明
```

## 使用方法
> 暂时只支持宏观研报生成，这里只是初级版本，后续高级版本还在本地优化。可以在src/workflows/macro_workflow0624.py中运行
