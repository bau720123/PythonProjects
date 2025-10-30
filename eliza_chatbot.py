import re

def eliza_response(user_input):
    # 去除大小寫與多餘空格，模擬 DOS 時代的簡單輸入處理
    user_input = user_input.lower().strip()
    
    # 定義模式匹配規則 (pattern matching)，確保字串正確關閉
    patterns = [
        (r'i feel (.*)', r'Why do you feel \1?'),
        (r'i am (.*)', r'Tell me more about being \1.'),
        (r'(.*)sad(.*)', r'Why does \1 make you sad?'),
        (r'(.*)happy(.*)', r"I'm glad to hear you're happy about \1!"),
        (r'(.*)taiwan semiconductor(.*)', r'Tell me more about your interest in Taiwan Semiconductor! Did you catch the recent 232 tariff news?'),
        (r'(.*)stock(.*)', r'Interesting! Are you thinking about stocks like TSMC? What’s your strategy?'),
        (r'(.*)regret(.*)', r'Regrets, huh? Like missing a TSMC trade at 1185? Tell me more.'),
        (r'(.*)', r'Tell me more about that.')  # 通用回應，模仿 Eliza 的萬用句
    ]
    
    # 遍歷模式，匹配用戶輸入
    for pattern, response in patterns:
        match = re.match(pattern, user_input)
        if match:
            # 替換關鍵詞，生成回應
            return response.replace('\\1', match.group(1)) if match.groups() else response
    
    return "Hmm, I'm not sure how to respond to that. Can you tell me more?"

def eliza_chatbot():
    print("Welcome to Eliza Chatbot (Inspired by DOS Era)! Type 'exit' to quit.")
    print("Talk to me like I'm your therapist or a friend curious about TSMC!")
    
    while True:
        user_input = input("You: ")
        if user_input.lower() == 'exit':
            print("Goodbye! Hope we can chat again, maybe about TSMC's next move!")
            break
        response = eliza_response(user_input)
        print(f"Eliza: {response}")

if __name__ == "__main__":
    eliza_chatbot()