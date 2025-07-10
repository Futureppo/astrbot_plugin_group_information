from typing import Any, Dict, List
import pandas as pd
from datetime import datetime
from io import BytesIO
import base64
from astrbot.api.star import Star, register, Context
from astrbot.api.event.filter import PermissionType
from astrbot.core.platform.message_type import MessageType
from astrbot.api import logger
from astrbot.api.event import filter
from astrbot.core.platform.sources.aiocqhttp.aiocqhttp_message_event import (
    AiocqhttpMessageEvent,
)


@register(
    "astrbot_plugin_group_information",
    "Futureppo",
    "导出群成员信息为Excel表格",
    "1.0.8",
    "https://github.com/Futureppo/astrbot_plugin_group_information",
)
class GroupInformationPlugin(Star):
    def __init__(self, context: Context):
        super().__init__(context)

    @staticmethod
    def _format_timestamp(timestamp):
        """格式化时间戳为可读时间，出错则用对应位的零代替"""
        if isinstance(timestamp, (int, float)) and timestamp >= 0:
            return datetime.fromtimestamp(float(timestamp)).strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        else:
            return "0000-00-00 00:00:00"

    @staticmethod
    def _clean_excel_invalid_chars(text):
        """清理Excel不支持的特殊字符"""
        if not isinstance(text, str):
            return text
        return "".join(
            char for char in text if ord(char) >= 32 and char not in "\x00\x01\x02\x03"
        )

    @filter.command("导出群数据")
    async def export_group_data(self, event: AiocqhttpMessageEvent, group_id: str = ''):
        """导出指定群聊成员信息到Excel文件"""
        if group_id:
            group_id = group_id.strip()
            if not group_id.isdigit():
                yield event.plain_result("请输入有效的群号")
                return
        else:
            group_id = event.get_group_id()
            if not group_id:
                yield event.plain_result("请在群聊中使用此命令或提供有效的群号")
                return
        try:
            client = event.bot

            try:
                is_group_member = await client.get_group_member_info(
                    group_id=int(group_id), user_id=int(event.get_sender_id()), no_cache=True
                )  # type: ignore
            except Exception as e:
                logger.error(f"获取群成员信息时出错: {e}")
                is_group_member = False
            if not is_group_member:
                yield event.plain_result("你不在该群聊中，无法导出数据")
                return
            # 获取群成员列表
            members: list[dict] = await client.get_group_member_list(
                group_id=int(group_id), no_cache=True
            )  # type: ignore
            # 处理成员数据
            processed_members = self._process_members(members)
            # 生成Excel文件
            file_content = self._generate_excel_file(
                processed_members, sheet_name=f"Group_{group_id}"
            )
            # 设置文件名
            file_name = f"群聊{group_id}的{len(processed_members)}名成员的数据.xlsx"
            # 上传文件
            message_type = 'group' if event.message_obj.type == MessageType.GROUP_MESSAGE else 'private'
            if message_type == 'group':
                group_or_user_id = event.get_group_id()
            else:
                group_or_user_id = event.get_sender_id()
            await self._upload_file(
                event, file_content, group_or_user_id, file_name, message_type=message_type
            )

        except Exception as e:
            logger.error(f"导出群数据时出错: {e}")
            yield event.plain_result(f"导出群数据时出错")

    @filter.permission_type(PermissionType.ADMIN)
    @filter.command("导出所有群数据")
    async def export_all_groups_data(self, event: AiocqhttpMessageEvent):
        """导出所有群的成员信息到多个sheet的Excel文件中"""

        client = event.bot
        group_list = await client.get_group_list(no_cache=True)
        yield event.plain_result(f"正在导出{len(group_list)}个群的数据...")
        try:
            # 创建Excel工作簿到内存中
            output_buffer = BytesIO()
            total_members = 0
            # 遍历所有群
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                for group in group_list:
                    group_id = group["group_id"]
                    group_name = group["group_name"]

                    try:
                        members: list[dict] = await client.get_group_member_list(
                            group_id=group_id, no_cache=True
                        )  # type: ignore
                        processed_members = self._process_members(members)
                        for member in processed_members:
                            member["group_name"] = self._clean_excel_invalid_chars(
                                group_name
                            )

                        df = pd.DataFrame(processed_members)
                        sheet_name = f"G{group_id}"[:30]  # Excel表名单字节限制
                        df.to_excel(writer, index=False, sheet_name=sheet_name)

                        total_members += len(processed_members)
                        logger.info(f"已导出{group_name}({group_id})的{len(processed_members)}名成员信息")

                    except Exception as e:
                        logger.error(f"处理群 {group_id} 时出错: {e}")
                        continue
            # 返回结果
            output_buffer.seek(0)
            file_content = output_buffer.getvalue()
            # 设置文件名
            file_name = (
                f"{len(group_list)}个群的{total_members}名成员的数据.xlsx"
            )
            # 上传文件
            message_type = 'group' if event.message_obj.type == MessageType.GROUP_MESSAGE else 'private'
            if message_type == 'group':
                group_or_user_id = event.get_group_id()
            else:
                group_or_user_id = event.get_sender_id()
            await self._upload_file(
                event, file_content, group_or_user_id, file_name, message_type=message_type
            )

        except Exception as e:
            logger.error(f"导出所有群数据时出错: {e}")
            yield event.plain_result("导出所有群数据时出错")

    def _process_members(self, members: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """处理成员数据，清理字段并格式化时间戳"""
        processed_members = []

        char_clean_fields = {"nickname", "card", "title"}
        timestamp_fields = {
            "join_time": None,
            "last_sent_time": None,
            "title_expire_time": 0,
            "shut_up_timestamp": 0,
        }

        for member in members:
            if not isinstance(member, dict):
                logger.warning("发现非字典类型的成员数据，已跳过")
                continue

            processed = {}

            for key, value in member.items():
                if key in char_clean_fields and isinstance(value, str):
                    processed[key] = self._clean_excel_invalid_chars(value)
                elif key in timestamp_fields:
                    processed[key] = value
                else:
                    processed[key] = value

            for key, default in timestamp_fields.items():
                raw_time = processed.get(key, default)
                processed[key] = self._format_timestamp(raw_time)

            processed_members.append(processed)

        return processed_members

    @staticmethod
    def _generate_excel_file(
        data: List[Dict[str, Any]], sheet_name: str = "Sheet1"
    ) -> bytes:
        """生成Excel文件"""
        df = pd.DataFrame(data)
        output_buffer = BytesIO()
        df.to_excel(
            output_buffer, index=False, engine="openpyxl", sheet_name=sheet_name
        )
        output_buffer.seek(0)
        return output_buffer.getvalue()

    @staticmethod
    async def _upload_file(
        event: AiocqhttpMessageEvent,
        file_content: bytes,
        user_or_group_id: str | int,
        file_name: str,
        message_type: str,
    ) -> bool:
        """上传文件到群组"""
        client = event.bot
        try:
            file_content_base64 = base64.b64encode(file_content).decode("utf-8")
            if message_type == 'group':
                await client.upload_group_file(
                    group_id=int(user_or_group_id),
                    file=f"base64://{file_content_base64}",
                    name=file_name,
                )
            else:
                await client.upload_private_file(
                    user_id=int(user_or_group_id),
                    file=f"base64://{file_content_base64}",
                    name=file_name,
                )
            logger.info(f"文件上传完成：{file_name}")
            return True
        except Exception as upload_e:
            logger.error(f"文件上传失败：{upload_e}")
            return False
