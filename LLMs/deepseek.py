# Please install OpenAI SDK first: `pip install openai`
from openai import OpenAI
# 创建 API 客户端
client = OpenAI(api_key="sk-xxxxxxxxxxxxxxxxxxx", base_url="https://api.deepseek.com")

# 调用 deepseek-chat 模型
response = client.chat.completions.create(
    model="deepseek-chat",
    messages=[
        {"role": "system", "content": "you are an expert in entity extraction."},
        {"role": "user", "content": "假如你是一个实体抽取的专家，现在我有以下实体类型，包含“电磁波类型”“长度”“宽度”等，请你从以下句子中抽取实体类型对应的实体，并以“实体类型:实体名称”的方式返回。句子为:我希望创建一个长500mm，宽度300mm，高度45mm的长方形电磁波Airbox; 特征阻抗是92ohm的表层微带线。基板的介电常数DK是3.6，介电损耗是0.002。输出:"},
    ],
    stream=False  # 设置为 True 可启用流式输出
)

# 输出响应内容
print(response.choices[0].message.content)