#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件重命名补丁脚本
将所有 {日期}-N9K&LINKAS ACL交叉检查.xlsx 重命名为 {日期}-跨平台N9K&LINKAS&OOB ACL交叉检查.xlsx
支持所有日期的文件重命名
"""

import os
import sys
import re

def rename_file_function():
    """重命名ACL交叉检查任务生成的文件


    将所有 {日期}-N9K&LINKAS ACL交叉检查.xlsx 重命名为
    {日期}-跨平台N9K&LINKAS&OOB ACL交叉检查.xlsx


    Returns:
        bool: 如果成功重命名至少一个文件则返回True，否则返回False
    """
    # 定义目录和文件名
    # 脚本在 v12/Patch 目录，目标文件在 v12/LOG/ACLCrossCheckTask/ 目录
    script_directory = os.path.dirname(os.path.abspath(__file__))
    v12_directory = os.path.dirname(script_directory)  # v12 目录
    target_directory = os.path.join(v12_directory, "LOG", "ACLCrossCheckTask")


    # 检查目标目录是否存在
    if not os.path.exists(target_directory):
        print(f"目标目录不存在: {target_directory}")
        return False


    # 匹配旧格式的文件名：{日期}-N9K&LINKAS ACL交叉检查.xlsx
    old_pattern = r"^(\d{8})-N9K&LINKAS ACL交叉检查\.xlsx$"


    # 查找所有匹配的文件
    all_files = os.listdir(target_directory)
    matched_files = []


    for filename in all_files:
        match = re.match(old_pattern, filename)
        if match:
            date = match.group(1)
            old_file_path = os.path.join(target_directory, filename)
            new_filename = f"{date}-跨平台N9K&LINKAS&OOB ACL交叉检查.xlsx"
            new_file_path = os.path.join(target_directory, new_filename)
            matched_files.append((old_file_path, new_file_path, filename, new_filename))


    if not matched_files:
        print("未找到需要重命名的文件")
        return False


    # 处理每个匹配的文件
    success_count = 0
    skip_count = 0
    error_counter = 0


    for old_file_path, new_file_path, old_filename, new_filename in matched_files:
        # 检查新文件是否已存在
        if os.path.exists(new_file_path):
            print(f"⚠ 目标文件已存在，跳过: {new_filename}")
            skip_count += 1
            continue


        # 执行重命名
        try:
            os.rename(old_file_path, new_file_path)
            print(f"✓ 文件重命名成功: {old_filename} -> {new_filename}")
            success_count += 1
        except Exception as e:
            print(f"✗ 文件重命名失败: {old_filename} - {e}")
            error_counter += 1


    # 输出统计信息
    print(f"\n重命名完成: 成功 {success_count} 个, 跳过 {skip_count} 个, 失败 {error_counter} 个")
    return success_count > 0

if __name__ == "__main__":
    success = rename_file_function()
    sys.exit(0 if success else 1)

