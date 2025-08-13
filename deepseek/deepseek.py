# Please install OpenAI SDK first: `pip install openai`
from openai import OpenAI
# 创建 API 客户端
client = OpenAI(api_key="sk-number", base_url="https://api.deepseek.com")

# 调用 deepseek-chat 模型
response = client.chat.completions.create(
    model="deepseek-chat",
    messages=[
        {"role": "system", "content": "You are a helpful assistant."},
        {"role": "user", "content": "我现在在和deepseek进行对话，但是现在很不方便，可以帮我生成一个UI界面吗？ 我在UI界面描述问题，你回答我问题也在UI界面。 希望这个界面是python代码"},
    ],
    stream=False  # 设置为 True 可启用流式输出
)

# 输出响应内容
print(response.choices[0].message.content)