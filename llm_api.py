from openai import OpenAI
import os
client = OpenAI(
    api_key=os.environ["DEEPSEEK_API_KEY"],
    base_url="https://api.deepseek.com/v1"
)
def call_llm_for_question(prompt, q_type):
    try:
        # 按题型分配合适的 temperature 与 max_tokens
        if q_type == 'choice':
            temperature = 0.7
            max_tokens = 512
        elif q_type == 'blank':
            temperature = 0.8  # 填空题也要稍微灵活一点点
            max_tokens = 768
        elif q_type == 'short':
            temperature = 1.0
            max_tokens = 1024
        elif q_type == 'code':
            temperature = 1.0
            max_tokens = 2048
        else:
            # 未知题型默认安全设置
            temperature = 0.7
            max_tokens = 512

        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个教育系统中的AI出题助手，返回清晰规范的问题格式"},
                {"role": "user", "content": prompt}
            ],
            temperature=temperature,
            max_tokens=max_tokens
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"[调用失败] {e}"

def call_llm_for_plan(prompt):
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一名教学辅助AI，擅长生成完整教学设计方案。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=2048
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"[调用失败] {e}"

def call_llm_for_feedback(prompt: str) -> str:
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个教育系统中的AI批改助手，专注于分析学生答案并给出错误定位与修正建议"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=1024
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"[调用失败] {e}"

def call_llm_for_analysis(prompt):
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个教学评估专家，擅长根据学生答题数据进行整体分析，总结班级掌握情况和教学建议。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.6,
            max_tokens=2048
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"[分析失败] {e}"

def call_llm_for_scoring(prompt):
    # 用于评估学生答案
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个教育评分系统，请对学生答案进行评分，输出一个0-100之间的整数表示得分。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=100
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"[调用失败] {e}"

def evaluate_student_answer(student_answer, reference_answer, question_text):
    # 封装prompt
    prompt = f"""请你对学生的作答进行评分：
题目：{question_text}
参考答案：{reference_answer}
学生作答：{student_answer}
请根据答案内容准确性给出一个整数分数（0~100），并只返回该整数，不要返回其它内容。"""
    score_text = call_llm_for_scoring(prompt)
    try:
        return int(score_text.strip().split()[0])
    except:
        return 0  # fallback

def call_llm_for_ask(prompt):
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个教学系统中的智能学习助手，回答学生的问题时要结合教学内容，回答清晰、准确、通俗易懂"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=1024
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"[调用失败] {e}"

def call_llm_for_realtime(prompt):
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "你是一个教育系统中的AI出题助手，返回清晰规范的问题格式，可以一次返回多题"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.8,  #稍微灵活一点
            max_tokens=2048   #支持多题输出
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"[调用失败] {e}"
