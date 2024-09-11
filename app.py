import openai
import gradio as gr
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")


# Load Excel data
def load_excel_data():
    file_path = r'GPTdata0911.xlsx'
    df = pd.read_excel(file_path)
    return df


# Load product data
df = load_excel_data()

# Initialize conversation and current step
conversation = []
current_step = "步驟零"

# Function to handle user input and generate recommendations
def query_chatgpt(user_input, state, email):
    global conversation, df, current_step

    # Define system prompt
    system_prompt = f"""
    提提供你一份「GPTdata0911」Excel表，其中每一項產品均有「產品名稱」、「公司名稱」、「公司地址」、「連絡電話」、「產品網址」等基本資訊，並描述產品「主要功能」和「使用方式」，以及「產品第一層分類」和「產品第二層分類」。
    你是智慧照顧產品推薦專家，你的任務是根據客戶的需求，從「GPTdata0911」Excel檔案中推薦合適的產品。每次推薦時，請依照以下步驟進行：
    步驟零：收到提示詞「開始新的推薦」時，請開始新的推薦流程，閱讀「GPTdata0911」Excel表中的信息。
    步驟一：客戶提出簡短需求，你初步判斷需求產品屬於哪一類別（來自Excel表中的「產品第一層分類」），並簡短解釋該分類意義後，詢問客戶確認是否是其要的類別。
    步驟二：客戶確認類別後，仔細閱讀該分類下所有產品的「主要功能」和「使用方式」，並提出可能的問題以幫助進一步篩選正確產品。   
    步驟三：統整客戶需求，確定其回應對應的是相關的「產品第二層分類」，並再次詢問客戶並且確認需求。
    步驟四：篩選出最合適的產品至少三項，提供產品名稱、產商攤位號、簡單描述、產品網頁和廠商聯繫資訊。
    步驟五：詢問客戶是否滿意，並且告知「如果滿意請於下方填寫您的電子郵件，將會把本次推薦的結果寄送給您」；如果不滿意，回到步驟二，重新詢問客戶需求和其對應類別，並且繼續推進步驟。
    Excel表內容如下：
    {df.to_string(index=False)}
    """

    # Add user input to conversation history
    conversation.append({"role": "user", "content": user_input})

    # Generate response using OpenAI API
    response = openai.ChatCompletion.create(
        model="gpt-4o-2024-08-06",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "system", "content": f"目前進行：{current_step}"},
            *conversation
        ]
    )

    # Get ChatGPT's reply
    reply = response['choices'][0]['message']['content']

    # Add assistant's reply to conversation
    conversation.append({"role": "assistant", "content": reply})

    # Update state
    state["recommendations"] = reply
    state["email_content"] = reply

    # Create conversation history
    conversation_history = [(conversation[i]['content'], conversation[i+1]['content']) for i in range(0, len(conversation) - 1, 2)]

    return conversation_history, state

# Gradio interface
def gradio_interface(user_input, email, state):
    if state is None:
        state = {"step": 0, "top_matches": None, "products_info": None, "recommendations": "", "email_content": "", "chat_history": []}
    return query_chatgpt(user_input, state, email)

# Gradio Blocks UI
with gr.Blocks(css="#logo img { box-shadow: none; border: none; }") as demo:
    with gr.Row():
        with gr.Column(scale=1):
            logo = gr.Image(r"GRC.png", elem_id="logo")
        with gr.Column(scale=6):
            gr.Markdown("# **智慧照顧產品推薦系統**")
        with gr.Column(scale=2):
            gr.Markdown("**免責聲明**：本系統應用ChatGPT進行智慧照顧產品推薦，提供之產品資訊僅供參考，使用者應自行前往各產品的官方網頁確認詳細資訊及最新規格。")
    
    # State management
    state = gr.State({"step": 0, "dialog_history": []})

    with gr.Row():
        with gr.Column(scale=12):
            chatbot = gr.Chatbot()
        with gr.Column(scale=4):
            user_input = gr.Textbox(label="您的訊息", placeholder="在此輸入...", lines=1, show_label=False)  # Single line input box
            email = gr.Textbox(label="電子郵件", placeholder="請在此輸入您的電子郵件信箱")
            send_email_btn = gr.Button("寄送郵件")
            # Add QR code under the email section
            qr_code = gr.Image(r"QRCode.png", label="掃描此QR Code填寫回饋表單", elem_id="qr_code", scale=4)  # Make sure to use appropriate file path
    
    # Interactions
    def interact(user_input, state, email):
        chat_history, state = query_chatgpt(user_input, state, email)
        return chat_history, state, ""

    def handle_send_email(email, state):
        return [("郵件已發送", "")]
    
    # Trigger on Enter key
    user_input.submit(interact, inputs=[user_input, state, email], outputs=[chatbot, state, user_input])
    send_email_btn.click(handle_send_email, inputs=[email, state], outputs=[chatbot])

demo.launch(share=True)
