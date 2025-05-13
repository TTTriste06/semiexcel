
def upload_to_github(file, path_in_repo, commit_message):
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{path_in_repo}"
    
    file.seek(0)  # 确保指针在开头
    file_content = file.read()
    encoded_content = base64.b64encode(file_content).decode('utf-8')

    # 先获取文件 SHA（如果存在）
    response = requests.get(api_url, headers={"Authorization": f"token {GITHUB_TOKEN}"})
    if response.status_code == 200:
        sha = response.json()['sha']
    else:
        sha = None

    # 构造 payload
    payload = {
        "message": commit_message,
        "content": encoded_content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha

    # 上传文件
    response = requests.put(api_url, json=payload, headers={"Authorization": f"token {GITHUB_TOKEN}"})

    # 结果反馈
    if response.status_code in [200, 201]:
        st.success(f"{path_in_repo} 上传成功！")
    else:
        st.error(f"上传失败: {response.json()}")
