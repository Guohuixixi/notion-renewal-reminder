#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
import smtplib
import ssl
import sys
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage


NOTION_TOKEN = os.environ["NOTION_TOKEN"]
DATABASE_ID = os.environ["NOTION_DATABASE_ID"]

SMTP_HOST = os.environ["SMTP_HOST"]
SMTP_PORT = int(os.environ["SMTP_PORT"])
SMTP_USER = os.environ["SMTP_USER"]
SMTP_PASSWORD = os.environ["SMTP_PASSWORD"]
MAIL_FROM = os.environ["MAIL_FROM"]
MAIL_TO_DEFAULT = os.environ["MAIL_TO"]

NOTION_VERSION = "2022-06-28"
TZ = timezone(timedelta(hours=8))

PROPERTY_ACCOUNT = "用户账号"
PROPERTY_BUYER_EMAIL = "买家邮箱"
PROPERTY_CONTACT_EMAIL = "联系邮箱"
PROPERTY_PLAN_TYPE = "套餐类型"
PROPERTY_MONTHS = "月数"
PROPERTY_START_DATE = "服务开始日"
PROPERTY_END_DATE = "服务到期日"
PROPERTY_RENEWED = "是否已续费"
PROPERTY_REMINDER_SENT = "提醒已发送"
PROPERTY_NOTIFY_EMAIL = "通知邮箱"
PROPERTY_NOTE = "备注"


def notion_request(method, path, body=None):
    if body is None:
        data = None
    else:
        data = json.dumps(body).encode("utf-8")

    request = urllib.request.Request(
        url=f"https://api.notion.com/v1{path}",
        data=data,
        method=method,
        headers={
            "Authorization": f"Bearer {NOTION_TOKEN}",
            "Notion-Version": NOTION_VERSION,
            "Content-Type": "application/json",
        },
    )

    try:
        with urllib.request.urlopen(request) as response:
            response_text = response.read().decode("utf-8")
            return json.loads(response_text)
    except urllib.error.HTTPError as error:
        error_text = error.read().decode("utf-8")
        print("Notion API 请求失败：", file=sys.stderr)
        print(error_text, file=sys.stderr)
        raise


def query_due_tomorrow():
    tomorrow = (datetime.now(TZ).date() + timedelta(days=1)).isoformat()

    print(f"今天日期：中国时间 {datetime.now(TZ).date().isoformat()}")
    print(f"本次查询的到期日：{tomorrow}")

    body = {
        "page_size": 100,
        "filter": {
            "and": [
                {
                    "property": PROPERTY_END_DATE,
                    "date": {
                        "equals": tomorrow
                    }
                },
                {
                    "property": PROPERTY_RENEWED,
                    "checkbox": {
                        "equals": False
                    }
                },
                {
                    "property": PROPERTY_REMINDER_SENT,
                    "checkbox": {
                        "equals": False
                    }
                }
            ]
        }
    }

    all_rows = []
    start_cursor = None

    while True:
        if start_cursor:
            body["start_cursor"] = start_cursor

        result = notion_request(
            method="POST",
            path=f"/databases/{DATABASE_ID}/query",
            body=body
        )

        rows = result.get("results", [])
        all_rows.extend(rows)

        has_more = result.get("has_more", False)
        start_cursor = result.get("next_cursor")

        if not has_more:
            break

    return all_rows


def get_plain_value(prop):
    if not prop:
        return ""

    prop_type = prop.get("type")

    if prop_type == "title":
        title_items = prop.get("title", [])
        return "".join(item.get("plain_text", "") for item in title_items)

    if prop_type == "rich_text":
        text_items = prop.get("rich_text", [])
        return "".join(item.get("plain_text", "") for item in text_items)

    if prop_type == "email":
        return prop.get("email") or ""

    if prop_type == "phone_number":
        return prop.get("phone_number") or ""

    if prop_type == "url":
        return prop.get("url") or ""

    if prop_type == "number":
        value = prop.get("number")
        if value is None:
            return ""
        return str(value)

    if prop_type == "select":
        select_value = prop.get("select")
        if not select_value:
            return ""
        return select_value.get("name", "")

    if prop_type == "status":
        status_value = prop.get("status")
        if not status_value:
            return ""
        return status_value.get("name", "")

    if prop_type == "date":
        date_value = prop.get("date")
        if not date_value:
            return ""
        return date_value.get("start", "")

    if prop_type == "checkbox":
        return "是" if prop.get("checkbox") else "否"

    return ""


def parse_email_list(email_text):
    if not email_text:
        return []

    normalized = (
        email_text
        .replace("，", ",")
        .replace("；", ",")
        .replace(";", ",")
    )

    emails = []

    for item in normalized.split(","):
        email = item.strip()
        if email:
            emails.append(email)

    return emails


def send_email(subject, body, to_addresses):
    if not to_addresses:
        raise ValueError("收件人邮箱为空，无法发送邮件")

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = MAIL_FROM
    message["To"] = ", ".join(to_addresses)
    message.set_content(body)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(message)


def mark_reminder_sent(page_id):
    body = {
        "properties": {
            PROPERTY_REMINDER_SENT: {
                "checkbox": True
            }
        }
    }

    notion_request(
        method="PATCH",
        path=f"/pages/{page_id}",
        body=body
    )


def build_email_content(row):
    props = row.get("properties", {})

    account = get_plain_value(props.get(PROPERTY_ACCOUNT)) or "未填写账号"
    buyer_email = get_plain_value(props.get(PROPERTY_BUYER_EMAIL))
    contact_email = get_plain_value(props.get(PROPERTY_CONTACT_EMAIL))
    plan_type = get_plain_value(props.get(PROPERTY_PLAN_TYPE))
    months = get_plain_value(props.get(PROPERTY_MONTHS))
    start_date = get_plain_value(props.get(PROPERTY_START_DATE))
    end_date = get_plain_value(props.get(PROPERTY_END_DATE))
    note = get_plain_value(props.get(PROPERTY_NOTE))
    notify_email = get_plain_value(props.get(PROPERTY_NOTIFY_EMAIL))

    if notify_email:
        to_addresses = parse_email_list(notify_email)
    else:
        to_addresses = parse_email_list(MAIL_TO_DEFAULT)

    page_url = row.get("url", "")

    subject = f"【续费提醒】{account} 将于 {end_date} 到期"

    body = f"""以下 Notion 商业版账号即将到期，请确认是否续费。

账号信息：
- 用户账号：{account}
- 套餐类型：{plan_type}
- 服务月数：{months}
- 服务开始日：{start_date}
- 服务到期日：{end_date}

联系信息：
- 买家邮箱：{buyer_email}
- 联系邮箱：{contact_email}

备注：
{note}

Notion 记录：
{page_url}

提醒规则：
这封邮件是因为该账号将在明天到期，并且 Notion 表格中：
- 是否已续费 = 未勾选
- 提醒已发送 = 未勾选

发送成功后，脚本会自动把「提醒已发送」勾选上，避免重复提醒。

--
此邮件由 GitHub Actions 自动发送。
"""

    return subject, body, to_addresses, account


def main():
    print("开始执行 Notion 商业版账号到期提醒脚本")

    rows = query_due_tomorrow()

    print(f"找到 {len(rows)} 条明天到期、未续费、未提醒的记录。")

    if not rows:
        print("没有需要提醒的记录，本次运行结束。")
        return

    success_count = 0
    failed_count = 0

    for row in rows:
        page_id = row.get("id")

        try:
            subject, body, to_addresses, account = build_email_content(row)

            print("----------------------------------------")
            print(f"准备发送提醒：{account}")
            print(f"收件人：{', '.join(to_addresses)}")

            send_email(
                subject=subject,
                body=body,
                to_addresses=to_addresses
            )

            mark_reminder_sent(page_id)

            success_count += 1
            print(f"发送成功：{account}")

        except Exception as error:
            failed_count += 1
            print("----------------------------------------", file=sys.stderr)
            print("发送失败：", file=sys.stderr)
            print(str(error), file=sys.stderr)

    print("----------------------------------------")
    print(f"执行完成。成功：{success_count} 条，失败：{failed_count} 条。")

    if failed_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
