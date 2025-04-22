import streamlit as st
import page1
import page2

st.set_page_config(page_title='投标文件修订系统',layout='wide')
st.sidebar.title('功能导航')
page=st.sidebar.radio('请选择功能',['暗标word格式调整','智能文档修订系统'])

if page=='暗标word格式调整':
    page1.app()
if page=='智能文档修订系统':
    page2.app()