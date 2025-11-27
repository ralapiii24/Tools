# 代码规范检查工具

# 导入标准库
import os
import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Tuple

# 导入第三方库
# (无第三方库依赖)

# 导入本地应用
# (无本地应用依赖)

# 代码规范检查器类：按照V11代码规范检查所有Python文件
class CodeStyleChecker:
    
    # 初始化代码规范检查器：设置检查规则和文件路径
    def __init__(self, root_dir: str = "."):
        self.ROOT_DIR = root_dir
        self.ISSUES = defaultdict(list)
        self.FILE_COUNT = 0
        self.ISSUE_COUNT = 0
        
        # 预期的导入分类
        self.EXPECTED_IMPORT_SECTIONS = [
            "# 导入标准库",
            "# 导入第三方库",
            "# 导入本地应用"
        ]
    
    # 获取所有Python文件：递归查找所有.py文件，排除__pycache__和隐藏目录
    def _get_python_files(self) -> List[str]:
        PYTHON_FILES = []
        for ROOT, DIRS, FILES in os.walk(self.ROOT_DIR):
            # 排除__pycache__和隐藏目录
            DIRS[:] = [D for D in DIRS if D != '__pycache__' and not D.startswith('.')]
            for FILE in FILES:
                if FILE.endswith('.py'):
                    FILE_PATH = os.path.join(ROOT, FILE)
                    PYTHON_FILES.append(FILE_PATH)
        return sorted(PYTHON_FILES)
    
    # 检查文件头部结构：检查模块说明和空行
    def _check_file_header(self, lines: List[str], file_path: str) -> None:
        if len(lines) == 0:
            self.ISSUES[file_path].append(("文件结构", "文件为空"))
            return
        
        # 检查第1行：模块说明或shebang
        if not lines[0].strip().startswith('#') and not lines[0].strip().startswith('#!/'):
            self.ISSUES[file_path].append(("文件结构", f"第1行：缺少模块说明注释（当前：{lines[0][:50]}）"))
        
        # 查找第一个模块说明注释行（跳过shebang和编码声明）
        module_comment_line = None
        for i, line in enumerate(lines[:5]):  # 只检查前5行
            stripped = line.strip()
            if stripped.startswith('#') and not stripped.startswith('#!/') and not stripped.startswith('# -*-'):
                module_comment_line = i
                break
        
        # 如果找到模块说明注释，检查其后是否有空行
        if module_comment_line is not None and module_comment_line + 1 < len(lines):
            if lines[module_comment_line + 1].strip() != '':
                self.ISSUES[file_path].append(("文件结构", f"第{module_comment_line + 2}行：模块说明后缺少空行"))
    
    # 检查导入分类：检查导入是否分为三类并正确标注
    def _check_import_sections(self, lines: List[str], file_path: str) -> None:
        IMPORT_LINES = []
        for LINE_INDEX, LINE in enumerate(lines, 1):
            if '# 导入' in LINE:
                IMPORT_LINES.append((LINE_INDEX, LINE.strip()))
        
        if not IMPORT_LINES:
            # __init__.py可能不需要导入分类
            if os.path.basename(file_path) != '__init__.py':
                self.ISSUES[file_path].append(("导入规范", "未找到导入分类注释"))
            return
        
        # 检查导入分类是否完整
        FOUND_SECTIONS = [LINE[1] for LINE in IMPORT_LINES]
        for EXPECTED in self.EXPECTED_IMPORT_SECTIONS:
            if EXPECTED not in FOUND_SECTIONS:
                # TaskBase.py可能不需要本地应用导入
                if EXPECTED == "# 导入本地应用" and os.path.basename(file_path) == "TaskBase.py":
                    continue
                # __init__.py可能不需要标准导入分类
                if os.path.basename(file_path) == '__init__.py':
                    continue
                self.ISSUES[file_path].append(("导入规范", f"缺少导入分类：{EXPECTED}"))
        
        # 检查导入分类顺序
        if len(IMPORT_LINES) >= 1:
            if IMPORT_LINES[0][1] != self.EXPECTED_IMPORT_SECTIONS[0]:
                self.ISSUES[file_path].append(("导入规范", f"第{IMPORT_LINES[0][0]}行：导入分类顺序错误，应首先出现'{self.EXPECTED_IMPORT_SECTIONS[0]}'"))
        
        if len(IMPORT_LINES) >= 2:
            if IMPORT_LINES[1][1] != self.EXPECTED_IMPORT_SECTIONS[1]:
                self.ISSUES[file_path].append(("导入规范", f"第{IMPORT_LINES[1][0]}行：导入分类顺序错误，应出现'{self.EXPECTED_IMPORT_SECTIONS[1]}'"))
        
        if len(IMPORT_LINES) >= 3:
            if IMPORT_LINES[2][1] != self.EXPECTED_IMPORT_SECTIONS[2]:
                self.ISSUES[file_path].append(("导入规范", f"第{IMPORT_LINES[2][0]}行：导入分类顺序错误，应出现'{self.EXPECTED_IMPORT_SECTIONS[2]}'"))
    
    # 检查docstring：禁止使用三引号注释
    def _check_docstring(self, lines: List[str], file_path: str) -> None:
        for LINE_INDEX, LINE in enumerate(lines, 1):
            if '"""' in LINE or "'''" in LINE:
                # 排除正则表达式的原始字符串（r"""...""" 或 r'''...'''）
                if LINE.strip().startswith('r"""') or LINE.strip().startswith("r'''"):
                    continue
                # 排除赋值语句中的正则表达式（如 pattern = re.compile(r"""...""")）
                if '= re.compile(' in LINE or '= re.' in LINE:
                    continue
                # 检查上下文，判断是否是函数/类的docstring
                if LINE_INDEX > 1:
                    PREV_LINE = lines[LINE_INDEX-2].strip() if LINE_INDEX > 1 else ''
                    if PREV_LINE.startswith('def ') or PREV_LINE.startswith('class '):
                        self.ISSUES[file_path].append(("注释规范", f"第{LINE_INDEX}行：禁止使用docstring，应使用单行注释"))
    
    # 检查类和函数注释：检查类和函数定义前是否有注释
    def _check_class_function_comments(self, lines: List[str], file_path: str) -> None:
        for LINE_INDEX, LINE in enumerate(lines, 1):
            STRIPPED = LINE.strip()
            if STRIPPED.startswith('class '):
                # 检查类前是否有注释（跳过@dataclass等装饰器）
                # 向前查找，跳过装饰器和空行，找到第一个非空行
                found_comment = False
                for CHECK_INDEX in range(LINE_INDEX - 2, -1, -1):
                    if CHECK_INDEX < 0:
                        break
                    CHECK_LINE = lines[CHECK_INDEX].strip()
                    if CHECK_LINE == '':
                        continue
                    if CHECK_LINE.startswith('#'):
                        found_comment = True
                        break
                    if CHECK_LINE.startswith('@'):
                        continue  # 跳过装饰器
                    break  # 遇到其他内容，停止查找
                if not found_comment:
                    self.ISSUES[file_path].append(("注释规范", f"第{LINE_INDEX}行：类定义前缺少注释"))
            elif STRIPPED.startswith('def ') and not STRIPPED.startswith('def _'):
                # 检查公共函数前是否有注释（跳过私有函数，后面单独检查）
                # 向前查找，跳过装饰器和空行，找到第一个非空行
                found_comment = False
                for CHECK_INDEX in range(LINE_INDEX - 2, -1, -1):
                    if CHECK_INDEX < 0:
                        break
                    CHECK_LINE = lines[CHECK_INDEX].strip()
                    if CHECK_LINE == '':
                        continue
                    if CHECK_LINE.startswith('#'):
                        found_comment = True
                        break
                    if CHECK_LINE.startswith('@'):
                        continue  # 跳过装饰器
                    break  # 遇到其他内容，停止查找
                if not found_comment:
                    self.ISSUES[file_path].append(("注释规范", f"第{LINE_INDEX}行：函数定义前缺少注释"))
            elif STRIPPED.startswith('def _'):
                # 检查私有函数前是否有注释
                # 向前查找，跳过装饰器和空行，找到第一个非空行
                found_comment = False
                for CHECK_INDEX in range(LINE_INDEX - 2, -1, -1):
                    if CHECK_INDEX < 0:
                        break
                    CHECK_LINE = lines[CHECK_INDEX].strip()
                    if CHECK_LINE == '':
                        continue
                    if CHECK_LINE.startswith('#'):
                        found_comment = True
                        break
                    if CHECK_LINE.startswith('@'):
                        continue  # 跳过装饰器
                    break  # 遇到其他内容，停止查找
                if not found_comment:
                    self.ISSUES[file_path].append(("注释规范", f"第{LINE_INDEX}行：私有函数定义前缺少注释"))
    
    # 检查单字母变量：禁止使用单字母循环变量
    def _check_single_letter_variables(self, lines: List[str], file_path: str) -> None:
        for LINE_INDEX, LINE in enumerate(lines, 1):
            # 跳过注释行
            if LINE.strip().startswith('#'):
                continue
            # 检查单字母循环变量（排除一些特殊情况）
            MATCHES = re.findall(r'\bfor\s+([a-z])\s+in\b', LINE)
            for VAR in MATCHES:
                # x, y 可能是坐标，暂时允许
                if VAR not in ['x', 'y']:
                    self.ISSUES[file_path].append(("命名规范", f"第{LINE_INDEX}行：禁止使用单字母循环变量 '{VAR}'"))
    
    # 检查连续空行：最多允许2个连续空行
    def _check_empty_lines(self, lines: List[str], file_path: str) -> None:
        EMPTY_COUNT = 0
        for LINE_INDEX, LINE in enumerate(lines, 1):
            if LINE.strip() == '':
                EMPTY_COUNT += 1
            else:
                if EMPTY_COUNT > 2:
                    self.ISSUES[file_path].append(("代码结构", f"第{LINE_INDEX-EMPTY_COUNT}行：连续空行{EMPTY_COUNT}个（最多允许2个）"))
                EMPTY_COUNT = 0
    
    # 检查单个文件：执行所有检查规则
    def _check_file(self, file_path: str) -> None:
        try:
            with open(file_path, 'r', encoding='utf-8') as FILE_HANDLE:
                LINES = FILE_HANDLE.readlines()
            
            # 移除行尾换行符，便于处理
            LINES = [LINE.rstrip('\n') for LINE in LINES]
            
            # 执行各项检查
            self._check_file_header(LINES, file_path)
            self._check_import_sections(LINES, file_path)
            self._check_docstring(LINES, file_path)
            self._check_class_function_comments(LINES, file_path)
            self._check_single_letter_variables(LINES, file_path)
            self._check_empty_lines(LINES, file_path)
            
        except Exception as ERROR:
            self.ISSUES[file_path].append(("文件读取", f"无法读取文件：{ERROR}"))
    
    # 运行检查：检查所有Python文件，返回问题字典
    def run_check(self) -> Dict[str, List[Tuple[str, str]]]:
        PYTHON_FILES = self._get_python_files()
        self.FILE_COUNT = len(PYTHON_FILES)
        
        for FILE_PATH in PYTHON_FILES:
            self._check_file(FILE_PATH)
        
        # 统计问题数量
        for ISSUES in self.ISSUES.values():
            self.ISSUE_COUNT += len(ISSUES)
        
        return dict(self.ISSUES)
    
    # 生成报告：生成格式化的检查报告
    def generate_report(self) -> str:
        REPORT_LINES = []
        REPORT_LINES.append("=" * 70)
        REPORT_LINES.append("代码规范检查报告")
        REPORT_LINES.append("=" * 70)
        REPORT_LINES.append(f"检查文件数：{self.FILE_COUNT}")
        REPORT_LINES.append(f"存在问题文件数：{len(self.ISSUES)}")
        REPORT_LINES.append(f"总问题数：{self.ISSUE_COUNT}")
        REPORT_LINES.append("")
        
        if not self.ISSUES:
            REPORT_LINES.append("✓ 所有文件检查通过！")
            REPORT_LINES.append("")
            REPORT_LINES.append("=" * 70)
            return "\n".join(REPORT_LINES)
        
        # 按类别统计
        CATEGORY_COUNT = defaultdict(int)
        for FILE_PATH, ISSUES in self.ISSUES.items():
            for CATEGORY, _ in ISSUES:
                CATEGORY_COUNT[CATEGORY] += 1
        
        REPORT_LINES.append("问题分类统计：")
        for CATEGORY, COUNT in sorted(CATEGORY_COUNT.items()):
            REPORT_LINES.append(f"  {CATEGORY}: {COUNT} 个问题")
        REPORT_LINES.append("")
        
        REPORT_LINES.append("详细问题列表：")
        REPORT_LINES.append("-" * 70)
        for FILE_PATH in sorted(self.ISSUES.keys()):
            REPORT_LINES.append(f"\n【{FILE_PATH}】")
            for CATEGORY, ISSUE in self.ISSUES[FILE_PATH]:
                REPORT_LINES.append(f"  [{CATEGORY}] {ISSUE}")
        
        REPORT_LINES.append("")
        REPORT_LINES.append("=" * 70)
        return "\n".join(REPORT_LINES)
    
    # 生成摘要报告：生成简要的问题摘要
    def generate_summary(self) -> str:
        if not self.ISSUES:
            return "✓ 所有文件检查通过！"
        
        # 按类别统计
        CATEGORY_COUNT = defaultdict(int)
        for FILE_PATH, ISSUES in self.ISSUES.items():
            for CATEGORY, _ in ISSUES:
                CATEGORY_COUNT[CATEGORY] += 1
        
        SUMMARY_LINES = []
        SUMMARY_LINES.append(f"检查结果：{len(self.ISSUES)}/{self.FILE_COUNT} 个文件存在问题")
        SUMMARY_LINES.append("")
        SUMMARY_LINES.append("问题分类：")
        for CATEGORY, COUNT in sorted(CATEGORY_COUNT.items()):
            SUMMARY_LINES.append(f"  {CATEGORY}: {COUNT} 个问题")
        
        return "\n".join(SUMMARY_LINES)

# 主函数：运行代码规范检查并输出报告
def main():
    CHECKER = CodeStyleChecker(".")
    CHECKER.run_check()
    
    # 输出报告
    REPORT = CHECKER.generate_report()
    print(REPORT)
    
    # 返回问题数量作为退出码（0表示无问题）
    return 0 if CHECKER.ISSUE_COUNT == 0 else 1

if __name__ == "__main__":
    exit(main())

