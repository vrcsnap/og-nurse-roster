import streamlit as st
import base64
import os

st.set_page_config(
    page_title="Duty Roster System",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 隱藏所有 Streamlit 標題、選單、頁尾與邊距，全螢幕居中呈現 GIF
st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container {
        padding: 0 !important;
        margin: 0 !important;
        max-width: 100% !important;
        height: 100vh;
        display: flex;
        justify-content: center;
        align-items: center;
        background-color: #000000;
    }
    .troll-container {
        display: flex;
        justify-content: center;
        align-items: center;
        width: 100vw;
        height: 100vh;
    }
    .troll-img {
        max-width: 90vw;
        max-height: 90vh;
        object-fit: contain;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

image_path = "dancing-dance.jpg"
gif_url = "https://media.giphy.com/media/Vuw9m5wXviFIQ/giphy.gif"

if os.path.exists(image_path):
    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode()
    fallback_src = f"data:image/jpeg;base64,{img_b64}"
else:
    fallback_src = gif_url

st.markdown(
    f"""
    <div class="troll-container">
        <img class="troll-img" src="{gif_url}" onerror="this.onerror=null; this.src='{fallback_src}';" alt="Troll" />
    </div>
    """,
    unsafe_allow_html=True
)
