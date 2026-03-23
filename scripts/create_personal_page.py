import pandas as pd
import json
import os
from datetime import datetime, timedelta
import shutil
import argparse
import jieba
from jieba import analyse
from wordcloud import WordCloud
from opencc import OpenCC
import requests
import time
import math
import sys

# 设置 stdout 编码为 UTF-8，支持 Emoji 输出
# 注意：在 Windows PowerShell 中可能会导致 I/O 错误，暂时注释掉
# if sys.platform == "win32":
#     import io
#     sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# import emoji
import pycountry
from emojiflags.lookup import lookup as flag
from multi_download import get_account_stat, get_online_stats_data
from common_tools import (
    db_path,
    read_db_table,
    insert_or_update_db,
    pic_to_webp,
    remove_blank_lines,
)
import pytz
import shutil

import re
from jinja2 import Template
import toml
import toml

BIN = os.path.dirname(os.path.realpath(__file__))
COOKIE_CONFIG_FILE = os.path.join(BIN, ".cookie_config.toml")

import os

config = toml.load("scripts/config.toml")
personal_page_link = config.get("url").get("personal_page_link")
# 优先从环境变量读取 Cookie，其次从 Cookie 配置文件
Cookie = os.environ.get("POSTCROSSING_COOKIE", "")
if not Cookie and os.path.exists(COOKIE_CONFIG_FILE):
    cookie_config = toml.load(COOKIE_CONFIG_FILE)
    Cookie = cookie_config.get("auth", {}).get("cookie", "")
pic_driver_path = config.get("url").get("pic_driver_path")
story_pic_link = config.get("url").get("story_pic_link")
story_pic_type = config.get("settings").get("story_pic_type")


def read_template_file():
    # 读取模板
    with open(
        os.path.join(BIN, f"../template/card_type.html"),
        "r",
        encoding="utf-8",
    ) as f:
        card_type_template = Template(f.read())

    with open(
        os.path.join(BIN, f"../template/信息汇总_index_template.txt"), "r", encoding="utf-8"
    ) as f:
        summary_template = Template(f.read())
    with open(
        os.path.join(BIN, f"../template/register_info_template.html"),
        "r",
        encoding="utf-8",
    ) as f:
        register_info_template = Template(f.read())
    # 读取年度汇总模板
    with open(
        os.path.join(BIN, f"../template/信息汇总_year_template.txt"), "r", encoding="utf-8"
    ) as f:
        year_summary_template = Template(f.read())
    return card_type_template, summary_template, register_info_template, year_summary_template


def update_sheet_data(excel_file):
    import warnings

    warnings.filterwarnings("ignore", category=FutureWarning)
    df = pd.read_excel(excel_file, na_filter=False, keep_default_na=False)
    df_json = df.to_dict(orient="records")

    def update_story_safe(key, item, existed_story):
        data = item.get(key) if item.get(key) else existed_story.get(key)
        return data

    default_item = {
        "card_id": "",
        "content_original": "",
        "content_cn": "",
        "comment_original": "",
        "comment_cn": "",
    }
    for item in df_json:
        existed_story = read_db_table(
            db_path, "postcard_story", {"card_id": item.get("id")}
        )
        if existed_story:
            default_item = existed_story[0]
        item_new = {
            "card_id": item.get("id"),
            "content_original": update_story_safe(
                "content_original", item, default_item
            ),
            "content_cn": update_story_safe("content_cn", item, default_item),
            "comment_original": update_story_safe(
                "comment_original", item, default_item
            ),
            "comment_cn": update_story_safe("comment_cn", item, default_item),
        }
        insert_or_update_db(db_path, "postcard_story", item_new)


def get_calendar_list():
    online_stats_data = get_online_stats_data(account)
    calendar_list = []

    for data in online_stats_data:
        timestamp = data[0]  # 获取时间戳
        date = datetime.fromtimestamp(timestamp)  # 将时间戳转换为日期格式
        year = date.strftime("%Y")  # 提取年份（YYYY）
        if year not in calendar_list:
            calendar_list.append(year)
    calendar_list = sorted(calendar_list, reverse=True)
    return calendar_list


def create_word_cloud(type, contents):
    keywords_old = []
    old_key_word_path = os.path.join(BIN, f"../output/keyword_old_{type}.txt")
    contents = contents.replace("nan", "")
    exclude_keywords = []  # 直接在这里指定排除的关键字
    if os.path.exists(old_key_word_path):

        with open(old_key_word_path, "r", encoding="utf-8") as f:
            keywords_old = [line.strip() for line in f.readlines()]
    # 转换为svg格式输出
    if type == "cn":

        path = cn_path_svg
        # 使用jieba的textrank功能提取关键词
        keywords = jieba.analyse.textrank(
            contents, topK=100, withWeight=False, allowPOS=("ns", "n", "vn", "v")
        )

        # 创建 OpenCC 对象，指定转换方式为简体字转繁体字
        converter = OpenCC("s2t.json")
        # 统计每个关键词出现的次数
        keyword_counts = {}
        for keyword in keywords:
            count = contents.count(keyword)
            keyword = converter.convert(keyword)  # 简体转繁体
            keyword_counts[keyword] = count
        # 创建一个WordCloud对象，并设置字体文件路径和轮廓图像
        wordcloud = WordCloud(
            width=1600, height=800, background_color="white", font_path=font_path
        )
        wordcloud.generate_from_frequencies(keyword_counts)
    else:
        path = en_path_svg
        wordcloud = WordCloud(
            width=1600,
            height=800,
            background_color="white",
            font_path=font_path,
            max_words=100,
        ).generate(contents)
        keywords = wordcloud.words_

    only_in_keywords = set(keywords) - set(keywords_old)
    only_in_keywords_old = set(keywords_old) - set(keywords)

    if not (only_in_keywords or only_in_keywords_old):
        print(f"keyword_{type}无更新，终止任务")
        return
    # 生成词云

    with open(old_key_word_path, "w", encoding="utf-8") as f:
        for keyword in keywords:
            f.write(f"{keyword}\n")
        print(f"\n✅ 已更新：{old_key_word_path}")
    svg_image = wordcloud.to_svg(embed_font=True)

    with open(path, "w+", encoding="UTF8") as f:
        f.write(svg_image)
        print(f"\n✅ 已保存至{path}")


def read_story_db(db_path):
    result_cn = ""
    result_en = ""
    content = read_db_table(db_path, "postcard_story")
    for item in content:

        map_info_data = read_db_table(
            db_path, "map_info", {"card_id": item.get("card_id")}
        )
        if map_info_data:
            item.update(map_info_data[0])
        content_original = item.get("content_original", "")
        content_cn = item.get("content_cn", "")
        comment_original = item.get("comment_original", "")
        comment_cn = item.get("comment_cn", "")
        
        # 清理标识文本
        content_original = clean_translation_markers(content_original)
        content_cn = clean_translation_markers(content_cn)
        comment_original = clean_translation_markers(comment_original)
        comment_cn = clean_translation_markers(comment_cn)
        
        data_en = f"{content_original}\n{comment_original}\n"
        data_cn = f"{content_cn}\n{comment_cn}\n"
        result_en += data_en
        result_cn += data_cn
    # print("result_en:", result_en)
    # print("result_cn:", result_cn)
    return result_cn.replace("None", ""), result_en.replace("None", "")


# 实时获取该账号所有sent、received的明信片列表，获取每个postcardID的详细数据


def calculate_days_difference(other_timestamp, sent_avg):

    current_timestamp = int(time.time())  # 当前时间戳
    traveled_days = math.floor((current_timestamp - other_timestamp) / 86400)
    if traveled_days >= 60:
        traveled_days_text = f'<span style="color: red;">{traveled_days}（过期）</span>'
    elif traveled_days > int(sent_avg):
        traveled_days_text = f'<span style="color: orange;">{traveled_days}</span>'
    else:
        traveled_days_text = f'<span style="color: green;">{traveled_days}</span>'
    return traveled_days, traveled_days_text


def get_traveling_id(account, Cookie):

    headers = {
        "authority": "www.postcrossing.com",
        "method": "GET",
        "path": f"/user/{account}/data/traveling",
        "scheme": "https",
        "accept": "application/json, text/javascript, */*; q=0.01",
        "accept-encoding": "gzip, deflate, br",
        "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "cookie": Cookie,
        "referer": f"https://www.postcrossing.com/user/{account}/traveling",
        "sec-ch-ua": '"Microsoft Edge";v="143", "Chromium";v="143", "Not A(Brand";v="24"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/143.0.0.0 Safari/537.36 Edg/143.0.0.0",
        "x-requested-with": "XMLHttpRequest",
    }

    # 使用 Session 来管理请求
    with requests.Session() as session:
        session.headers.update(headers)
        url = f"https://www.postcrossing.com/user/{account}/data/traveling"
        response = session.get(url)
        response.raise_for_status()  # 如果状态码不是 200，会抛出 HTTPError

        if response.status_code == 200:
            # 检查是否需要手动解压
            content = response.content
            json_str = content.decode("utf-8")  # 或者 'latin-1' 如果 utf-8 不行
            response = json.loads(json_str)
            # print("response:", response)

    traveling_count = len(response)
    content = sorted(response, key=lambda x: x[7])
    new_data = []

    def get_local_date(country_code, timestamp):
        # 根据国家二简码获取时区
        timezone = pytz.country_timezones.get(country_code)
        if timezone:
            timezone = pytz.timezone(timezone[0])
        else:
            return "Invalid country code"
        # 将时间戳转换为datetime对象
        dt = datetime.fromtimestamp(timestamp)
        # 将datetime对象转换为当地时区的时间
        local_dt = dt.astimezone(timezone)
        # 格式化日期为"%Y/%m/%d %H:%M"的字符串
        formatted_date = local_dt.strftime("%Y/%m/%d %H:%M")
        return formatted_date

    extra_info = []
    for i, stats in enumerate(content):
        sent_avg = 0
        baseurl = "https://www.postcrossing.com"
        country_stats_data = read_db_table(
            db_path, "country_stats", {"country_code": stats[3]}
        )
        country_list = read_db_table(
            db_path, "country_list", {"country_code": stats[3]}
        )

        if country_stats_data:
            sent_avg = country_stats_data[0].get("sent_avg", 0)
        if not sent_avg:
            sent_avg = 0
        traveled_days, traveled_days_text = calculate_days_difference(
            stats[4], sent_avg
        )

        item = {
            "card_id": f"<a href='{baseurl}/travelingpostcard/{stats[0]}' target='_blank'>{stats[0]}</a>",
            "sender": f"<a href='{baseurl}/user/{stats[1]}' target='_blank'>{stats[1]}</a>",
            "country_name": f"{country_list[0].get("country_name")} {country_list[0].get("country_name_emoji")}",
            "sent_local_date": get_local_date(stats[0][0:2], stats[4]),
            "distance": f'{format(stats[6], ",")}',
            "traveled_days": traveled_days,
            "traveled_days_text": traveled_days_text,
            "sent_avg": f"{sent_avg}",
        }
        extra_info.append(item)
    # expired_count = sum(1 for item in extra_info if "过期" in item["traveling_days"])
    expired_count = sum(1 for item in extra_info if int(item["traveled_days"]) > 60)
    html_content = card_type_template.render(
        card_type="traveling", content=extra_info, baseurl=baseurl
    )

    with open(f"./output/traveling.html", "w", encoding="utf-8") as file:
        file.write(html_content)
    return traveling_count, expired_count


def get_HTML_table(card_type, table_name):
    content = read_db_table(db_path, table_name, {"card_type": card_type})

    baseurl = "https://www.postcrossing.com"
    for i, stats in enumerate(content):
        # print("stats:", stats)
        country_key = "sent" if card_type == "received" else "received"
        country_stats_data = read_db_table(
            db_path, "country_stats", {"name": stats.get(f"{country_key}_country")}
        )
        if country_stats_data:
            stats.update(country_stats_data[0])
        stats["distance"] = format(stats.get("distance"), ",")

    title_name = read_db_table(db_path, "title_info", {"card_type": card_type})[0]
    content = sorted(content, key=lambda x: x["received_date_local"], reverse=True)
    html_content = card_type_template.render(
        card_type=card_type, content=content, title_name=title_name, baseurl=baseurl
    )

    with open(f"./output/{card_type}.html", "w", encoding="utf-8") as file:
        file.write(html_content)


def get_postcard_limit(sent_num):
    """
    https://www.postcrossing.com/help
    """
    if sent_num < 5:
        limit = 5
    elif sent_num < 35:
        limit = 6 + (sent_num - 5) // 10
    elif sent_num < 50:
        limit = 9
    else:
        limit = 10 + (sent_num - 50) // 50

    return min(limit, 100)


def create_register_info():
    """
    生成register_info.html
    """

    def get_user_sheet(table_name):
        content = read_db_table(db_path, table_name)
        country_count = len(content)
        content = sorted(content, key=lambda x: x["name"])
        html_content = card_type_template.render(
            card_type=table_name,
            content=content,
        )
        with open(f"./output/{table_name}.html", "w", encoding="utf-8") as file:
            file.write(html_content)
        return country_count

    countryNum = get_user_sheet("country_stats")
    countries = f"{countryNum}/248 [{round(countryNum/248*100,2)}%]"
    traveling_num, expired_num = get_traveling_id(account, Cookie)

    # 创建HTML内容
    item = read_db_table(db_path, "user_summary")[0]
    item["sent_distance"] = format(int(item.get("sent_distance")), ",")
    item["received_distance"] = format(int(item.get("received_distance")), ",")

    limit_num = get_postcard_limit(int(item.get("sent_postcard_num")))

    traveling = f"{traveling_num} [在途：{traveling_num-expired_num} | 过期：{expired_num} | 还可寄：{limit_num-traveling_num+expired_num}]"
    item.update(
        {
            "countries": countries,
            "traveling": traveling,
        }
    )
    html_content = register_info_template.render(item=item)

    # 写入HTML文件
    with open("./output/register_info.html", "w", encoding="utf-8") as file:
        file.write(html_content)


def get_card_type_data(card_type):
    """获取按年份分组的明信片数据"""
    data_list = read_db_table(db_path, "map_info", {"card_type": card_type})
    new_list = []
    for item in data_list:
        # 关联country_stats表数据
        country_stats_data = read_db_table(
            db_path,
            "country_stats",
            {
                "name": (
                    item.get("sent_country")
                    if card_type == "received"
                    else item.get("received_country")
                )
            },
        )
        if country_stats_data:
            item.update(country_stats_data[0])

        # 处理经纬度
        from_coor = json.loads(item.get("from_coor"))
        to_coor = json.loads(item.get("to_coor"))

        from_coor0 = from_coor[0] if from_coor else ""
        from_coor1 = from_coor[1] if from_coor else ""

        to_coor0 = to_coor[0] if to_coor[0] else ""
        to_coor1 = to_coor[1] if to_coor[1] else ""

        item.update(
            {
                "from_coor0": from_coor0,
                "from_coor1": from_coor1,
                "to_coor0": to_coor0,
                "to_coor1": to_coor1,
            }
        )
        item["distance"] = format(item.get("distance"), ",")

        # 关联postcard_story数据
        postcard_story = read_db_table(
            db_path, "postcard_story", {"card_id": item.get("card_id")}
        )
        if postcard_story:
            story_data = postcard_story[0]
            # 清理邮件回复中的多余空行（移除所有空行，只保留非空行）
            for key in ["comment_original", "comment_cn", "content_original", "content_cn"]:
                if story_data.get(key):
                    story_data[key] = remove_blank_lines(story_data[key])
            item.update(story_data)

        # 关联gallery_info数据
        gallery_info = read_db_table(
            db_path, "gallery_info", {"card_id": item.get("card_id")}
        )
        if gallery_info:
            item.update(gallery_info[0])

        if any(
            [
                item.get("content_original"),
                item.get("content_cn"),
                item.get("comment_original"),
                item.get("comment_cn"),
            ]
        ):
            new_list.append(item)

    content = {}
    for item in new_list:
        received_year = item["received_date"].split("/")[0]
        if received_year not in content:
            content[received_year] = []
        content[received_year].append(item)
    for year in content:
        content[year] = sorted(
            content[year], key=lambda x: x["received_date"], reverse=True
        )
    content = dict(sorted(content.items(), key=lambda x: x[0], reverse=True))

    return content, len(new_list)


def clean_translation_markers(text):
    """清理翻译和识别工具的标识文本"""
    if not text:
        return text
    
    import re
    
    # 定义需要过滤的标识文本模式（支持多行匹配）
    patterns = [
        # [由xxx] 格式的标识
        r"\[由imap_tools提取\]",
        r"\[由Gemini[^\]]*?\]",
        r"\[由[^\]]+?\]",
        # 无括号格式的标识
        r"由imap_tools提取",
        r"由Gemini\s+gemini\s+gemini\s+flash\s+flash\s+lite\s+lite\s+preview\s+preview\s+识别",
        r"由Gemini\s+gemini[\s\S]*?preview\s+识别",
        r"由Gemini[\s\S]*?识别",
        r"gemini\s+flash",
        r"flash\s+lite",
        r"lite\s+preview",
        r"preview\s+识别",
        r"^识别\s*",
        r"\s+识别\s*",
    ]
    
    # 逐个使用正则替换标识文本
    for pattern in patterns:
        text = re.sub(pattern, "", text, flags=re.MULTILINE | re.IGNORECASE)
    
    # 清理多余的空行（复用 remove_blank_lines）
    return remove_blank_lines(text)


def read_story_db_by_year(db_path, year):
    """按年份读取明信片故事数据"""
    result_cn = ""
    result_en = ""
    content = read_db_table(db_path, "postcard_story")
    for item in content:
        map_info_data = read_db_table(
            db_path, "map_info", {"card_id": item.get("card_id")}
        )
        if map_info_data:
            item.update(map_info_data[0])
            # 检查年份是否匹配
            received_date = item.get("received_date", "")
            if received_date and received_date.startswith(str(year)):
                content_original = item.get("content_original", "")
                content_cn = item.get("content_cn", "")
                comment_original = item.get("comment_original", "")
                comment_cn = item.get("comment_cn", "")
                
                # 清理标识文本
                content_original = clean_translation_markers(content_original)
                content_cn = clean_translation_markers(content_cn)
                comment_original = clean_translation_markers(comment_original)
                comment_cn = clean_translation_markers(comment_cn)
                
                data_en = f"{content_original}\n{comment_original}\n"
                data_cn = f"{content_cn}\n{comment_cn}\n"
                result_en += data_en
                result_cn += data_cn
    return result_cn.replace("None", ""), result_en.replace("None", "")


def create_summary_text():
    """
    生成信息汇总.md（索引页）和按年份拆分的信息汇总_YYYY.md
    """
    stat, content_raw, card_types = get_account_stat(account, Cookie)

    user_summary = read_db_table(db_path, "user_summary")[0]

    title_all = ""
    for card_type in card_types:
        title_info = read_db_table(db_path, "title_info", {"card_type": card_type})[0]
        title_name = title_info.get("title_name")
        title_all += f"#### [{title_name}](/{nick_name}/postcrossing/{card_type})\n\n"
        title_final = f"{title_all}"

    calendar_list = get_calendar_list()
    comment_list, comment_num = get_card_type_data("sent")
    story_list, story_num = get_card_type_data("received")
    
    # 替换 URL 中的 {{repo}} 占位符
    story_pic_link_replaced = story_pic_link.replace("{{repo}}", repo)
    pic_driver_path_replaced = pic_driver_path.replace("{{repo}}", repo)
    
    # 准备年份统计数据
    year_stats = {}
    for year in story_list.keys():
        year_stats[year] = {
            "story_count": len(story_list[year]),
            "comment_count": len(comment_list.get(year, []))
        }
    
    # 生成索引页（信息汇总.md）- 只包含公共部分
    index_data = summary_template.render(
        account=account,
        pic_driver_path=pic_driver_path_replaced,
        story_pic_link=story_pic_link_replaced,
        nick_name=nick_name,
        user_summary=user_summary,
        story_pic_type=story_pic_type,
        personal_page_link=personal_page_link,
        title_final=title_final,
        calendar_list=calendar_list,
        repo=repo,
        year_list=list(story_list.keys()),  # 年份列表用于导航
        year_stats=year_stats,  # 年份统计数据
    )
    
    with open(f"./gallery/信息汇总.md", "w", encoding="utf-8") as f:
        f.write(index_data)
    print(f"✅ 已生成信息汇总.md（索引页）")
    
    # 同步到博客目录（仅当目录存在时）
    blog_dir = r"D:\web\Blog\src\Arthur\Postcrossing"
    blog_path = os.path.join(blog_dir, "信息汇总.md")
    
    if os.path.exists(blog_dir):
        with open(blog_path, "w", encoding="utf-8") as f:
            f.write(index_data)
        print(f"✅ 已同步到博客: {blog_path}")
    
    # 计算年份对应的 order（年份越新，order 越小，排在越前面）
    year_list = list(story_list.keys())
    max_year = max(int(y) for y in year_list) if year_list else 2026
    
    # 生成每年的汇总页
    for year in story_list.keys():
        year_story_data = {year: story_list[year]}
        year_comment_data = {year: comment_list.get(year, [])}
        year_story_num = len(story_list[year])
        year_comment_num = len(comment_list.get(year, []))
        
        # 计算 order：最新年份 order=1，次年 order=2，以此类推
        order = max_year - int(year) + 1
        
        year_data = year_summary_template.render(
            account=account,
            pic_driver_path=pic_driver_path_replaced,
            story_pic_link=story_pic_link_replaced,
            nick_name=nick_name,
            user_summary=user_summary,
            story_pic_type=story_pic_type,
            personal_page_link=personal_page_link,
            title_final=title_final,
            calendar_list=calendar_list,
            repo=repo,
            year=year,
            story_list=year_story_data,
            story_num=year_story_num,
            comment_list=year_comment_data,
            comment_num=year_comment_num,
            order=order,
        )
        
        # 创建 gallery/各年详情/ 目录
        gallery_year_dir = f"./gallery/各年详情"
        os.makedirs(gallery_year_dir, exist_ok=True)
        
        year_file_path = f"{gallery_year_dir}/各年详情_{year}.md"
        with open(year_file_path, "w", encoding="utf-8") as f:
            f.write(year_data)
        print(f"✅ 已生成各年详情_{year}.md")
        
        # 同步到博客目录（仅当目录存在时）
        blog_year_dir = os.path.join(blog_dir, "各年详情")
        if os.path.exists(blog_dir):
            os.makedirs(blog_year_dir, exist_ok=True)
            blog_year_path = os.path.join(blog_year_dir, f"各年详情_{year}.md")
            with open(blog_year_path, "w", encoding="utf-8") as f:
                f.write(year_data)
            print(f"✅ 已同步到博客: {blog_year_path}")
        
        # 为每年生成独立的词云
        year_result_cn, year_result_en = read_story_db_by_year(db_path, year)
        if year_result_cn.strip():
            create_word_cloud_for_year("cn", year_result_cn, year)
        if year_result_en.strip():
            create_word_cloud_for_year("en", year_result_en, year)
    
    print(f"————————————————————")


def create_word_cloud_for_year(type, contents, year):
    """为特定年份生成词云"""
    keywords_old = []
    old_key_word_path = os.path.join(BIN, f"../output/keyword_old_{type}_{year}.txt")
    contents = contents.replace("nan", "")
    
    if os.path.exists(old_key_word_path):
        with open(old_key_word_path, "r", encoding="utf-8") as f:
            keywords_old = [line.strip() for line in f.readlines()]
    
    # 设置输出路径
    if type == "cn":
        path = os.path.join(BIN, f"../output/postcrossing_cn_{year}.svg")
        # 使用jieba的textrank功能提取关键词
        keywords = jieba.analyse.textrank(
            contents, topK=100, withWeight=False, allowPOS=("ns", "n", "vn", "v")
        )
        # 创建 OpenCC 对象，指定转换方式为简体字转繁体字
        converter = OpenCC("s2t.json")
        # 统计每个关键词出现的次数
        keyword_counts = {}
        for keyword in keywords:
            count = contents.count(keyword)
            keyword = converter.convert(keyword)  # 简体转繁体
            keyword_counts[keyword] = count
        # 创建一个WordCloud对象，并设置字体文件路径和轮廓图像
        wordcloud = WordCloud(
            width=1600, height=800, background_color="white", font_path=font_path
        )
        wordcloud.generate_from_frequencies(keyword_counts)
    else:
        path = os.path.join(BIN, f"../output/postcrossing_en_{year}.svg")
        wordcloud = WordCloud(
            width=1600,
            height=800,
            background_color="white",
            font_path=font_path,
            max_words=100,
        ).generate(contents)
        keywords = wordcloud.words_

    only_in_keywords = set(keywords) - set(keywords_old)
    only_in_keywords_old = set(keywords_old) - set(keywords)

    if not (only_in_keywords or only_in_keywords_old):
        print(f"keyword_{type}_{year}无更新，跳过")
        return
    
    # 生成词云
    with open(old_key_word_path, "w", encoding="utf-8") as f:
        for keyword in keywords:
            f.write(f"{keyword}\n")
        print(f"\n✅ 已更新：{old_key_word_path}")
    
    svg_image = wordcloud.to_svg(embed_font=True)
    with open(path, "w+", encoding="UTF8") as f:
        f.write(svg_image)
        print(f"\n✅ 已保存至{path}")


if __name__ == "__main__":
    card_type_template, summary_template, register_info_template, year_summary_template = read_template_file()

    # nick_name = data["nick_name"]

    # 创建 ArgumentParser 对象
    parser = argparse.ArgumentParser()
    parser.add_argument("account", help="输入account")
    parser.add_argument("nick_name", help="输入nickName")
    parser.add_argument("repo", help="输入repo1")
    options = parser.parse_args()

    account = options.account
    nick_name = options.nick_name
    repo = options.repo

    font_path = "./scripts/font.otf"
    cn_path_svg = "./output/postcrossing_cn.svg"
    en_path_svg = "./output/postcrossing_en.svg"
    excel_file = "./template/postcard_story.xlsx"

    pic_to_webp("./template/rawPic", "./template/content")
    excel_file = "./template/postcard_story.xlsx"
    update_sheet_data(excel_file)
    get_HTML_table("sent", "map_info")
    get_HTML_table("received", "map_info")
    create_register_info()

    create_summary_text()

    # 生成词云
    result_cn, result_en = read_story_db(db_path)
    create_word_cloud("cn", result_cn)
    create_word_cloud("en", result_en)
