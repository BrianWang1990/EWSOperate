#! /usr/bin/python
# -*- coding: utf-8 -*-
import base64
import datetime
import json
import logging

import pytz
from exchangelib import (  # Definition,
    DELEGATE,
    IMPERSONATION,
    Account,
    Configuration,
    Credentials,
    DLMailbox,
    EWSDate,
    EWSDateTime,
    EWSTimeZone,
    ExtendedProperty,
    FaultTolerance,
    Mailbox,
    Message,
    HTMLBody, FileAttachment)
from exchangelib.folders import FolderCollection, Messages

from it_platform.ews.constant import FOLDER_NAME

# 步骤1. 前端输入发件人邮箱地址和主题(可模糊搜索)
# 步骤2. 检索发件人的已发送(sent)文件夹 或是 所有文件夹？？？
# 步骤3. 前端可输入任一个收件人的邮箱地址和主题(可模糊搜索)
# 步骤4. 将上面3个步骤检索出的邮件，以列表形式展示到前端，供用户最终选择要删除的邮件(展示字段: message_id, subject, text_body的部分内容)
# 步骤5. 获取选中要删除的邮件的message_id
# 步骤6. 根据message_id 获取到收件人列表(display_to) 和 抄送列表(display_cc) 和 密送列表(无法获取？？？)
# 步骤7. 判断步骤6中列表里，是否有邮件组，如果有邮件组，需查找邮件组成员，拼接到步骤6后面
# 步骤8. 对列表进行去重  set 操作
# 步骤9. 遍历所有人的收件箱(inbox)是否有步骤5中的message_id，如果查到，则删除，并计数；如果未查到，则加到 未查到人员的数组里(not_in_inbox_list)
# 步骤10. 遍历步骤9中not_in_inbox_list 中的人员邮箱，非inbox文件夹中是否有步骤5中的message_id，如果查到，则删除，并计数，从not_in_inbox_list中删除
# 步骤11. 将没在步骤9和步骤10中查询到message_id的邮箱统计，输出前端页面
# 步骤12. 输出前端页面，共发送给了N个人，共删除M个人邮箱的相关记录，N-M个邮箱未查询到，未删除


# 声明一个有 ApplicationImpersonation 权限账号 模拟登录


class EWS:
    def __init__(self, user, pwd, endpoint):
        self.user = user
        self.credentials = Credentials(user, pwd)
        # c = Configuration(server="xxxxxx", credentials=credentials)
        # service_endpoint 是为了兼容http，如果服务器是https,可直接使用server="xxxxxxx"
        self.c = Configuration(
            service_endpoint=endpoint,
            retry_policy=FaultTolerance(max_wait=3600),
            credentials=self.credentials,
        )
        self.tz = pytz.timezone("Asia/Shanghai")
        self.email_addresses = []

    # 模拟用户登录 account_upn 格式为 "xxx@.demo.com"
    def _make_impersonation(self, account_upn):
        impersonal_account = Account(
            account_upn,
            credentials=self.credentials,
            config=self.c,
            access_type=IMPERSONATION,
        )
        return impersonal_account

    # # # 根据邮件组名称，查询组内成员列表
    # def _get_list_by_group(self, upn, group_address):
    #     impersonal_account = self._make_impersonation(upn)
    #     # email_addresses = []
    #     for mailbox in impersonal_account.protocol.expand_dl(
    #         DLMailbox(email_address=f"{group_address}", mailbox_type="PublicDL")
    #     ):
    #         if mailbox.mailbox_type == "PublicDL":
    #             self._get_list_by_group(upn, mailbox.email_address)
    #         self.email_addresses.append(mailbox.email_address)
    #     return self.email_addresses

    # 遍历所有文件夹，查找message_id
    def get_mail_by_message_id(self, upn, message_id):
        impersonal_account = self._make_impersonation(upn)
        all_folders_info = FolderCollection(
            account=impersonal_account,
            folders=(
                f
                for f in impersonal_account.root.walk()
                if f.CONTAINER_CLASS == "IPF.Note"
            ),
        )
        result = all_folders_info.filter(message_id=message_id)
        # for item in result:
        #     print(item.subject)
        #     print(item.message_id)
        #     print(item.display_to)
        return result

    # 通过邮件主题、时间范围查询要删除的邮件
    def get_message_id_by_subject(
        self, upn, folder, subject, start_time=None, end_time=None
    ):
        impersonal_account = self._make_impersonation(upn)
        # 转换成ews 时间
        if str(folder).lower() == FOLDER_NAME.INBOX:
            messages = impersonal_account.inbox.filter(subject__contains=f"{subject}")
            if start_time:
                start_time_ews = EWSDateTime.from_datetime(
                    self.tz.localize(datetime.datetime.fromisoformat(start_time))
                )
                messages = messages.filter(datetime_received__gte=start_time_ews)
            if end_time:
                end_time_ews = EWSDateTime.from_datetime(
                    self.tz.localize(datetime.datetime.fromisoformat(end_time))
                )
                messages = messages.filter(datetime_received__lte=end_time_ews)
        elif str(folder).lower() == FOLDER_NAME.SENT:
            messages = impersonal_account.sent.filter(subject__contains=f"{subject}")
            if start_time:
                start_time_ews = EWSDateTime.from_datetime(
                    self.tz.localize(datetime.datetime.fromisoformat(start_time))
                )
                messages = messages.filter(datetime_sent__gte=start_time_ews)
            if end_time:
                end_time_ews = EWSDateTime.from_datetime(
                    self.tz.localize(datetime.datetime.fromisoformat(end_time))
                )
                messages = messages.filter(datetime_sent__lte=end_time_ews)
        return messages.all()

    # 根据messge_id 删除邮件
    def delete_message_by_message_id(self, upn, message_id):
        impersonal_account = self._make_impersonation(upn)
        inbox = impersonal_account.inbox.filter(message_id=message_id)
        sent = impersonal_account.sent.filter(message_id=message_id)  # 新增，避免收件人或收件人群组包含发件人自己的情况
        is_delete = False
        message = None
        # 优先遍历收件箱和已发送邮件，没在收件箱和已发送的再遍历所有文件夹
        if inbox.exists():
            for item in inbox:
                message = item  # 保存副本，为了返回结果
                item.delete()
                logging.info(
                    f" delete mail success, mail's owner is {impersonal_account}"
                )
                is_delete = True
        if sent.exists():
            for item in sent:
                message = item  # 保存副本，为了返回结果
                item.delete()
                logging.info(
                    f" delete mail success, mail's owner is {impersonal_account}"
                )
                is_delete = True
        if not inbox.exists() and not sent.exists():  # 既没在收件箱也没在已发送
            #  遍历所有文件夹
            all_folders_info = FolderCollection(
                account=impersonal_account,
                folders=(
                    f
                    for f in impersonal_account.root.walk()
                    if f.CONTAINER_CLASS == "IPF.Note"
                ),
            )
            all_folders = all_folders_info.filter(message_id=message_id)
            for item in all_folders:
                message = item  # 保存副本，为了返回结果
                item.delete()
                logging.info(
                    f" delete mail success, mail's owner is {impersonal_account}"
                )
                is_delete = True
        return is_delete, message

    # # 删除邮件主入口
    # def delete_message(
    #     self, need_delete_upns_list, need_delete_groups_list, message_id
    # ):
    #     deleted_count = 0
    #     need_delete = 0
    #     for upn in need_delete_upns_list:
    #         is_delete = self.delete_message_by_message_id(upn, message_id)
    #         need_delete += 1
    #         deleted_count += 1 if is_delete else 0
    #     for group in need_delete_groups_list:
    #         upns = self._get_list_by_group(self.user, group)
    #         for upn in upns:
    #             is_delete = self.delete_message_by_message_id(upn, message_id)
    #             need_delete += 1
    #             deleted_count += 1 if is_delete else 0
    #     return need_delete, deleted_count

    # 根据邮件组名称，查询组内成员列表
    def get_group_members(self, group_address):
        impersonal_account = self._make_impersonation(self.user)
        for mailbox in impersonal_account.protocol.expand_dl(
            DLMailbox(email_address=f"{group_address}", mailbox_type="PublicDL")
        ):
            if mailbox.mailbox_type == "PublicDL":  # 邮件组嵌套着邮件组 判断
                self.get_group_members(mailbox.email_address)
            if not mailbox.mailbox_type == "PublicDL":   # 只有当邮箱账号不是group类型时候，才加到待删除列表
                self.email_addresses.append(mailbox.email_address)
        return self.email_addresses

    # 发送邮件
    def send_message(self, upn, subject, body, to_addresses=[], cc_addresses=[], bcc_addresses=[], pics=None, attachments=None):
        is_send = False
        message = ""
        try:
            impersonal_account = self._make_impersonation(upn)
            to_recipients, cc_recipients, bcc_recipients = [], [], []
            if not isinstance(to_addresses, list):
                to_addresses = eval(to_addresses)
            if not isinstance(cc_addresses, list):
                cc_addresses = eval(cc_addresses)
            if not isinstance(bcc_addresses, list):
                bcc_addresses = eval(bcc_addresses)
            for recipient in to_addresses:
                to_recipients.append(Mailbox(email_address=recipient))
            for cc in cc_addresses:
                cc_recipients.append(Mailbox(email_address=cc))
            for bcc in bcc_addresses:
                bcc_recipients.append(Mailbox(email_address=bcc))
            # Create message
            m = Message(account=impersonal_account,
                        folder=impersonal_account.sent,
                        subject=subject,
                        body=HTMLBody(body),
                        to_recipients=to_recipients,
                        cc_recipients=cc_recipients,
                        bcc_recipients=bcc_recipients)
            # attach files
            for attachment_name, attachment_content in attachments or []:
                base64_bytes = attachment_content.encode("ascii")
                sample_string_bytes = base64.b64decode(base64_bytes)
                file = FileAttachment(name=attachment_name, content=sample_string_bytes)
                m.attach(file)
            # attach images   需要配合html使用
            for pic_name, pic_content in pics or []:
                # pic_content = unicode(pic_content, "utf-8")
                base64_bytes = pic_content.encode("ascii")
                sample_string_bytes = base64.b64decode(base64_bytes)
                file = FileAttachment(name=pic_name, content=sample_string_bytes, is_inline=True,  content_type='image/png', content_id=pic_name)
                m.attach(file)
            is_send = True
            message = f"发送成功！pics is {pics}"
            m.send_and_save()
            # return is_send, message
        except Exception as e:
            message = f"发送失败，error msg is {str(e)}"
        return is_send, message


def get_ews(request=None, tenant=None):
    if not tenant and request:
        tenant = request.tenant if hasattr(request, "tenant") else request.user.tenant
    # tenant = request.tenant if hasattr(request, "tenant") else request.user.tenant
    ews_user = tenant.ews_conf["ews_user"]
    ews_pwd = tenant.ews_conf["ews_pwd"]
    ews_endpoint = tenant.ews_conf["ews_endpoint"]
    ews = EWS(ews_user, ews_pwd, ews_endpoint)
    return ews
