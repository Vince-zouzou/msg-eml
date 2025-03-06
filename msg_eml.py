import streamlit as st
import extract_msg
from email.message import EmailMessage
from email.policy import default
import mimetypes
import io
import zipfile

def convert_msg_to_eml(msg_file):
    # 打开并解析 .msg 文件
    msg = extract_msg.Message(msg_file)

    # 创建一个新的 EmailMessage 对象
    eml = EmailMessage(policy=default)

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
        attachment_data = attachment.data
        attachment_name = attachment.longFilename or attachment.shortFilename or 'attachment'
        mime_type, _ = mimetypes.guess_type(attachment_name)
        if mime_type is None:
            mime_type = 'application/octet-stream'
        maintype, subtype = mime_type.split('/', 1)
        eml.add_attachment(attachment_data, maintype=maintype, subtype=subtype, filename=attachment_name)

    # 将 eml 转换为字符串并编码为字节
    eml_content = eml.as_string().encode('utf-8')

    # 关闭 msg 对象
    msg.close()

    return eml_content

# 创建 ZIP 文件的函数
def create_zip(files_data):
    # 在内存中创建 ZIP 文件
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, content in files_data.items():
            zip_file.writestr(filename, content)
    
    # 将缓冲区指针移到开头
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# Streamlit 界面
st.sidebar.title("MSG-TO-EML")
files = st.file_uploader("Upload MSG file", type=["msg"], accept_multiple_files=True)

if st.button("Convert"):
    if files:
        with st.spinner('Converting...'):
            # 存储所有转换后的文件内容
            converted_files = {}
            for uploaded_file in files:
                # 转换文件
                eml_content = convert_msg_to_eml(uploaded_file)
                # 使用上传文件名生成 .eml 文件名
                base_name = uploaded_file.name.rsplit('.', 1)[0]
                eml_filename = f"{base_name}_converted.eml"
                converted_files[eml_filename] = eml_content
                st.write(f"已转换: {uploaded_file.name} -> {eml_filename}")

            # 创建 ZIP 文件
            zip_data = create_zip(converted_files)

            # 提供 ZIP 文件下载
            st.download_button(
                label="Download All as ZIP",
                data=zip_data,
                file_name="converted_emails.zip",
                mime="application/zip"
            )
            st.success("所有文件已转换并打包，点击上方按钮下载 ZIP 文件！")
    else:
        st.warning("请先上传至少一个 .msg 文件！")
