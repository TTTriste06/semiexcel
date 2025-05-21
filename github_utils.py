import base64
import requests
import streamlit as st
from urllib.parse import quote


# GitHub 配置
GITHUB_TOKEN_KEY = "GITHUB_TOKEN"  # secrets.toml 中的密钥名
REPO_NAME = "TTTriste06/semiment"
BRANCH = "main"


def upload_to_github(file_obj, filename):
    """
    上传文件对象到 GitHub 仓库（filename 应为已编码的 safe_name）
    """
    token = st.secrets[GITHUB_TOKEN_KEY]
    url = f"https://api.github.com/repos/{REPO_NAME}/contents/{filename}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    file_obj.seek(0)
    content = base64.b64encode(file_obj.read()).decode("utf-8")
    file_obj.seek(0)

    # 获取 SHA（如果文件已存在）
    sha = None
    get_resp = requests.get(url, headers=headers)
    if get_resp.status_code == 200:
        sha = get_resp.json().get("sha")

    payload = {
        "message": f"upload {filename}",
        "content": content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha

    put_resp = requests.put(url, headers=headers, json=payload)
    if put_resp.status_code not in [200, 201]:
        st.error(f"❌ GitHub 上传失败（状态码 {put_resp.status_code}）")
        st.code(put_resp.text)
        st.code(f"📝 Payload:\n{payload}")
        raise RuntimeError("GitHub 上传失败")
    else:
        st.success(f"✅ 成功上传文件至 GitHub：{filename}")


def download_from_github(filename):
    """
    从 GitHub 仓库下载指定文件内容（以二进制返回）
    - filename: 仓库中保存的文件名
    - 返回: bytes 内容（可用于 pd.read_excel(BytesIO(...))）
    """
    token = st.secrets[GITHUB_TOKEN_KEY]
    safe_filename = quote(filename)

    url = f"https://api.github.com/repos/{REPO_NAME}/contents/{safe_filename}?ref={BRANCH}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        json_resp = response.json()
        return base64.b64decode(json_resp["content"])
    else:
        raise FileNotFoundError(f"❌ GitHub 上找不到文件：{filename} (HTTP {response.status_code})")

