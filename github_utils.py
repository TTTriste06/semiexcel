import requests
import base64
import streamlit as st
import pandas as pd
from io import BytesIO

from config import GITHUB_TOKEN_KEY, REPO_NAME, BRANCH

GITHUB_TOKEN = st.secrets[GITHUB_TOKEN_KEY]


def upload_to_github(file, path_in_repo, commit_message):
    """
    将文件上传到 GitHub 仓库指定位置。

    参数：
    - file: BytesIO 文件对象或上传的 file-like 对象
    - path_in_repo: 仓库内路径（包括文件名）
    - commit_message: 提交信息
    """
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{path_in_repo}"

    file.seek(0)  # 确保指针在开头
    file_content = file.read()
    encoded_content = base64.b64encode(file_content).decode('utf-8')

    # 获取现有文件的 SHA（如果存在）
    response = requests.get(api_url, headers={"Authorization": f"token {GITHUB_TOKEN}"})
    if response.status_code == 200:
        sha = response.json().get('sha')
    else:
        sha = None

    # 构造提交 payload
    payload = {
        "message": commit_message,
        "content": encoded_content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha

    # PUT 请求上传文件
    response = requests.put(api_url, json=payload, headers={"Authorization": f"token {GITHUB_TOKEN}"})

    # 提示用户结果
    if response.status_code in [200, 201]:
        st.success(f"✅ {path_in_repo} 上传成功！")
    else:
        st.error(f"❌ 上传失败：{response.status_code} - {response.json().get('message', '未知错误')}")

def download_excel_from_repo(path_in_repo):
    """
    从 GitHub 仓库下载指定路径的文件，并返回 BytesIO 对象。

    参数：
    - path_in_repo: 仓库中的文件路径（包括文件名）

    返回：
    - BytesIO 对象（可用于 pd.read_excel）
    """
    raw_url = f"https://raw.githubusercontent.com/{REPO_NAME}/{BRANCH}/{path_in_repo}"
    response = requests.get(raw_url, headers={"Authorization": f"token {GITHUB_TOKEN}"})

    if response.status_code == 200:
        return BytesIO(response.content)
    else:
        st.error(f"❌ 下载失败：{response.status_code} - {response.json().get('message', '未知错误')}")
        return None
