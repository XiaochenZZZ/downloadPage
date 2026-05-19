#!/usr/bin/env python3
"""
沈阳市市场监督管理局 - 食品安全抽检信息通告附件下载脚本（交互式年月选择）
运行后在终端输入起止年月即可自动下载对应 xls 文件，文件保存在脚本同目录下。
"""

import os
import re
import time
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# ==================== 固定配置 ====================
BASE_URL = "https://scj.shenyang.gov.cn/zwgk/fdzdgknr/spypaq/"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://scj.shenyang.gov.cn/",
}
REQUEST_DELAY = 1.5  # 请求间隔（秒）

# 文件保存目录：脚本所在目录
SAVE_DIR = os.path.dirname(os.path.abspath(__file__)) or "."
# =================================================


def input_year_month(prompt):
    """交互式输入年月，返回 (year, month)"""
    while True:
        s = input(prompt).strip()
        match = re.match(r"(\d{4})[/\-.]?(\d{1,2})", s)
        if match:
            year, month = int(match.group(1)), int(match.group(2))
            if 2000 <= year <= 2100 and 1 <= month <= 12:
                return year, month
        print(" ❌ 格式错误，请输入如 2025/03 或 2025-3 的形式。")


def get_total_pages(first_page_url, session):
    """从首页源码中提取总页数"""
    try:
        resp = session.get(first_page_url, timeout=15)
        resp.encoding = "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        for script in soup.find_all("script"):
            if script.string and "countPage" in script.string:
                match = re.search(r"var countPage = (\d+)", script.string)
                if match:
                    return int(match.group(1))
        pagebox = soup.find("div", class_="pagebox")
        if pagebox:
            text = pagebox.get_text()
            match = re.search(r"共(\d+)页", text)
            if match:
                return int(match.group(1))
        print("⚠️ 未能获取总页数，默认设为1")
        return 1
    except Exception as e:
        print(f"❌ 获取总页数失败: {e}")
        return 1


def extract_articles_from_page(page_url, session):
    """从列表页提取所有文章标题和链接"""
    articles = []
    try:
        resp = session.get(page_url, timeout=15)
        resp.encoding = "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        ul = soup.find("ul", class_="xxgk_rul")
        if not ul:
            return articles
        for li in ul.find_all("li"):
            a_tag = li.find("a")
            if not a_tag:
                continue
            title = a_tag.get("title", "").strip()
            if not title:
                title = a_tag.get_text(strip=True)
            href = a_tag.get("href", "")
            if not href:
                continue
            full_url = urljoin(page_url, href)
            articles.append({"title": title, "url": full_url})
    except Exception as e:
        print(f"  ❌ 解析列表页失败: {e}")
    return articles


def extract_date_from_title(title):
    """从通告标题中提取日期"""
    pattern = r"(\d{4})年(\d{1,2})月(\d{1,2})日"
    match = re.search(pattern, title)
    if match:
        return int(match.group(1)), int(match.group(2)), int(match.group(3))
    return None, None, None


def download_xls_file(file_url, save_path, session):
    """下载单个 xls 文件"""
    if os.path.exists(save_path):
        print(f"    ⏭️  已存在，跳过: {os.path.basename(save_path)}")
        return True
    try:
        print(f"    📥 正在下载: {os.path.basename(save_path)}")
        resp = session.get(file_url, timeout=60, stream=True)
        resp.raise_for_status()
        with open(save_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        print(f"    ✅ 下载成功: {os.path.basename(save_path)}")
        return True
    except Exception as e:
        print(f"    ❌ 下载失败: {e} (URL: {file_url})")
        return False


def find_xls_links(article_url, session):
    """从文章详情页查找合格/不合格 xls 文件链接"""
    xls_dict = {}
    try:
        resp = session.get(article_url, timeout=15)
        resp.encoding = "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        for a_tag in soup.find_all("a", attrs={"appendix": "true"}):
            title_attr = a_tag.get("title", "")
            href = a_tag.get("href", "")
            if not href:
                continue
            full_url = urljoin(article_url, href)
            if "食品抽检合格信息" in title_attr:
                xls_dict["合格"] = full_url
            elif "食品抽检不合格信息" in title_attr:
                xls_dict["不合格"] = full_url
    except Exception as e:
        print(f"    ❌ 解析文章页面失败: {e}")
    return xls_dict


# ==================== 主程序 ====================
if __name__ == "__main__":
    print("=" * 60)
    print("沈阳市市场监督管理局 - 食品安全抽检信息通告下载工具")
    print("=" * 60)

    # 交互式输入年月范围
    print("\n请输入要下载通告的起止年月：")
    YEAR_FROM, MONTH_FROM = input_year_month("起始年月（如 2024/1 或 2024-01）: ")
    YEAR_TO, MONTH_TO = input_year_month("终止年月（如 2026/12）: ")

    print(f"\n筛选范围: {YEAR_FROM}-{MONTH_FROM:02d} 至 {YEAR_TO}-{MONTH_TO:02d}")
    print(f"文件保存目录: {os.path.abspath(SAVE_DIR)}")

    # 创建保存目录（其实即当前目录，一般已存在）
    os.makedirs(SAVE_DIR, exist_ok=True)

    # 初始化 requests 会话
    session = requests.Session()
    session.headers.update(HEADERS)

    # 第一步：获取总页数
    print("\n📊 正在获取栏目总页数...")
    total_pages = get_total_pages(BASE_URL + "index.html", session)
    print(f"✅ 共 {total_pages} 页")

    target_articles = []

    # 第二步：遍历所有分页，收集匹配的通告
    print("\n🔍 正在扫描所有页面，查找符合条件的通告...")
    for page_num in range(total_pages):
        if page_num == 0:
            page_url = BASE_URL + "index.html"
        else:
            page_url = BASE_URL + f"index_{page_num}.html"

        print(f"  📄 正在扫描第 {page_num + 1}/{total_pages} 页...")
        articles = extract_articles_from_page(page_url, session)
        print(f"    发现 {len(articles)} 篇文章")

        for art in articles:
            if "食品安全抽检信息通告" not in art["title"]:
                continue
            year, month, day = extract_date_from_title(art["title"])
            if year is None:
                # 备用：从 URL 中提取日期
                url_match = re.search(r"/(\d{6})/t(\d{8})_", art["url"])
                if url_match:
                    year = int(url_match.group(2)[:4])
                    month = int(url_match.group(2)[4:6])
                    day = int(url_match.group(2)[6:8])
            if year is None:
                print(f"    ⚠️ 无法提取日期，跳过: {art['title']}")
                continue
            if (YEAR_FROM, MONTH_FROM) <= (year, month) <= (YEAR_TO, MONTH_TO):
                target_articles.append({
                    "title": art["title"],
                    "url": art["url"],
                    "year": year,
                    "month": month,
                    "day": day,
                })

        time.sleep(REQUEST_DELAY)

    print(f"\n🎯 共找到 {len(target_articles)} 篇符合条件的通告")

    if not target_articles:
        print("❌ 未找到任何符合条件的通告，请检查筛选条件。")
        exit(0)

    # 第三步：逐篇下载附件
    print("\n📥 开始下载附件...")
    downloaded_count = 0
    for idx, art in enumerate(target_articles, 1):
        date_str = f"{art['year']}-{art['month']:02d}-{art['day']:02d}"
        print(f"\n  [{idx}/{len(target_articles)}] {date_str} | {art['title']}")
        print(f"      URL: {art['url']}")

        xls_links = find_xls_links(art["url"], session)
        if not xls_links:
            print(f"    ⚠️ 未找到任何 xls 文件链接")
            continue

        for file_type, file_url in xls_links.items():
            if file_type == "合格":
                save_name = f"食品抽检合格信息_{date_str}.xls"
            else:
                save_name = f"食品抽检不合格信息_{date_str}.xls"
            save_path = os.path.join(SAVE_DIR, save_name)
            if download_xls_file(file_url, save_path, session):
                downloaded_count += 1

        time.sleep(REQUEST_DELAY)

    print(f"\n{'=' * 60}")
    print(f"✨ 任务完成！共下载 {downloaded_count} 个文件")
    print(f"📂 文件保存在: {os.path.abspath(SAVE_DIR)}")
    print(f"{'=' * 60}")