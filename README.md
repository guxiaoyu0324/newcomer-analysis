# 新人数据自动化分析系统

## 本地运行

1. 安装Python 3.8+
2. 安装依赖：
```bash
pip install -r requirements.txt
```
3. 运行应用：
```bash
streamlit run app.py
```
4. 浏览器自动打开 http://localhost:8501

## 部署到 Streamlit Cloud（推荐）

### 步骤：

1. **创建GitHub仓库**
   - 将 `webapp/` 文件夹中的所有文件推送到GitHub
   - 文件结构：
     ```
     webapp/
     ├── app.py
     ├── requirements.txt
     └── README.md
     ```

2. **访问 Streamlit Cloud**
   - 打开 https://streamlit.io/cloud
   - 用GitHub账号登录

3. **部署应用**
   - 点击 "New app"
   - 选择你的GitHub仓库
   - 主文件路径：`app.py`
   - Python版本：3.11
   - 点击 "Deploy"

4. **分享链接**
   - 部署完成后会生成一个URL
   - 格式：`https://your-app-name.streamlit.app`
   - 将这个链接分享给同事即可使用

## 使用说明

1. 将周会数据文件夹打包成zip（包含`本周新增数据/`、`累计相关数据/`、`往期数据底表/`三个子文件夹）
2. 在网页上上传zip文件
3. 输入本周周期（如"0618-0624"）
4. 点击运行，等待处理完成
5. 查看表格和图表，点击下载按钮保存文件

## 生成的文件

- 新人指标监控全量表.xls
- 本周新人情况一览表.xls
- 试录题合格率变化趋势-新.xls + .png
- 面试通过率趋势-新.xls + .png
- 面试通过-新人总名单-分层版.xls
