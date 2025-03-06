import streamlit as st
import extract_msg
from email.message import EmailMessage
from email.policy import default
import mimetypes
from email import policy
import streamlit as st

def convert_msg_to_eml(msg_path, eml_path):
    # 打开并解析 .msg 文件
    msg = extract_msg.Message(msg_path)

    # 创建一个新的 EmailMessage 对象
    eml = EmailMessage(policy=policy.default)

    # 从 msg.header 中复制关键头部信息，清理换行符和回车符
    essential_headers = ['From', 'To', 'Subject', 'Date', 'Message-ID']
    for key in essential_headers:
        if key in msg.header:
            value = msg.header[key]
            if value:
                value = value.replace('\n', ' ').replace('\r', ' ')
                eml[key] = value

    # 处理邮件正文
    if msg.htmlBody and msg.body:  # 如果同时有 HTML 和纯文本正文
        eml.add_alternative(msg.body, subtype='plain')
        eml.add_alternative(msg.htmlBody.decode('utf-8'), subtype='html')
    elif msg.htmlBody:  # 只有 HTML 正文
        eml.set_content(msg.htmlBody.decode('utf-8'), subtype='html')
    elif msg.body:  # 只有纯文本正文
        eml.set_content(msg.body, subtype='plain')

    # 处理附件（如果有）
    for attachment in msg.attachments:
        # 获取附件内容和文件名
        attachment_data = attachment.data
        attachment_name = attachment.longFilename or attachment.shortFilename or 'attachment'

        # 猜测 MIME 类型
        mime_type, _ = mimetypes.guess_type(attachment_name)
        if mime_type is None:
            mime_type = 'application/octet-stream'
        maintype, subtype = mime_type.split('/', 1)

        # 添加附件到 eml
        eml.add_attachment(attachment_data, maintype=maintype, subtype=subtype, filename=attachment_name)

    # 保存为 .eml 文件
    with open(eml_path, 'w', encoding='utf-8') as eml_file:
        eml_file.write(eml.as_string())

    # 关闭 msg 对象
    msg.close()



st.sidebar.title("MSG-TO-EML")
files = st.file_uploader("Upload MSG file", type=["msg"],accept_multiple_files=True)
eml_file = st.sidebar.text_input("Output EML file name", "output.eml")
if st.button("Convert"): 
    eml = [eml_file + '_{}.eml'.format(i) for i in range(len(files))]
    with st.spinner('Loading...'):
        print("Converting...--------------____________--------------------->")  
        for i,j in zip(files,eml):
            print("Converting...--------------____________--------------------->")  
            convert_msg_to_eml(i, j)
            print(i.name)
        #print(eml_file)
        #convert_msg_to_eml(input_msg_file, output_eml_file)


