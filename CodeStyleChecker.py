#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
代码规范检查工具（遵循 PEP 8 和 Pyright 标准）
检查项目中的代码规范，包括：
- 文件夹命名规范（全大写，项目特殊约定）
- 文件名命名规范（驼峰命名，PEP 8 标准）
- 类名命名规范（驼峰命名 PascalCase，PEP 8 标准）
- 函数名命名规范（小写下划线 snake_case，PEP 8 标准）
- 私有函数命名规范（单下划线 + 小写下划线，PEP 8 标准）
- 变量名命名规范（普通变量小写下划线，常量全大写，PEP 8 标准）
- 注释规范（文件顶部模块说明、函数前注释）
- 导入顺序规范（标准库/第三方库/本地应用，PEP 8 标准）
- 导入分组空行检查（导入组之间应有空行，PEP 8 标准）
- 空行规范（最多连续2个空行，PEP 8 标准）
- 行长度检查（建议不超过99字符，最多120字符，PEP 8 标准）
- 缩进检查（使用4个空格，不使用Tab，PEP 8 标准）
- 尾随空格检查（行尾不应有空格，代码整洁）
- 文件末尾换行符检查（文件末尾应有换行符，PEP 8 标准）
- 文档字符串检查（公共函数和类应有docstring，PEP 8 标准）
- 异常处理检查（避免bare except，PEP 8 标准）
- TODO/FIXME注释检查（提醒开发者处理待办事项）
"""

import os
import re
import ast
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from collections import defaultdict

# 标准库列表（常见标准库）
STANDARD_LIBRARIES = {
    'os', 'sys', 're', 'json', 'time', 'datetime', 'pathlib', 'typing',
    'collections', 'dataclasses', 'enum', 'base64', 'subprocess', 'traceback',
    'io', 'locale', 'socket', 'ipaddress', 'functools', 'itertools'
}

# 需要检查的目录
CHECK_DIRS = ['v12']
# 需要忽略的目录
IGNORE_DIRS = {'__pycache__', '.git', 'node_modules', '.pytest_cache', '.mypy_cache'}
# 需要忽略的文件
IGNORE_FILES = {'.pyc', '.pyo', '.pyd', '.so', '.dll', '.dylib'}

class CodeStyleChecker:
    """代码规范检查器"""
    
    def __init__(self, root_dir: str = "."):
        self.root_dir = Path(root_dir)
        self.errors = []
        self.warnings = []
        self.stats = defaultdict(int)
        
    def check_all(self) -> Tuple[List[str], List[str], Dict[str, int]]:
        """执行所有检查"""
        self.errors = []
        self.warnings = []
        self.stats = defaultdict(int)
        
        for check_dir in CHECK_DIRS:
            dir_path = self.root_dir / check_dir
            if dir_path.exists():
                self._check_directory_structure(dir_path)
                self._check_python_files(dir_path)
        
        return self.errors, self.warnings, dict(self.stats)
    
    def _check_directory_structure(self, dir_path: Path):
        """检查目录结构命名规范"""
        for item in dir_path.rglob('*'):
            if item.is_dir():
                # 跳过忽略的目录
                if any(ignore in item.parts for ignore in IGNORE_DIRS):
                    continue
                
                # 检查目录名是否全大写（允许数字和下划线）
                dir_name = item.name
                if dir_name and not re.match(r'^[A-Z0-9_]+$', dir_name):
                    # 排除一些特殊目录（如 Patch）
                    if dir_name not in {'Patch'}:
                        self.warnings.append(f"目录命名不规范: {item.relative_to(self.root_dir)} (应为全大写)")
                        self.stats['dir_warnings'] += 1
    
    def _check_python_files(self, dir_path: Path):
        """检查Python文件"""
        for py_file in dir_path.rglob('*.py'):
            # 跳过忽略的目录
            if any(ignore in py_file.parts for ignore in IGNORE_DIRS):
                continue
            
            self._check_filename(py_file)
            self._check_file_content(py_file)
    
    def _check_filename(self, file_path: Path):
        """检查文件名命名规范（驼峰命名）"""
        filename = file_path.stem  # 不含扩展名
        
        # 特殊文件名例外
        if filename in {'__init__', '__main__', 'Main'}:
            return
        
        # Patch文件例外
        if 'Patch' in file_path.parts:
            return
        
        # 检查是否为驼峰命名（首字母大写，后续单词首字母大写）
        if not re.match(r'^[A-Z][a-zA-Z0-9]*$', filename):
            # 允许全大写（如 CONFIG）
            if not re.match(r'^[A-Z_]+$', filename):
                self.warnings.append(f"文件名命名不规范: {file_path.relative_to(self.root_dir)} (应为驼峰命名)")
                self.stats['filename_warnings'] += 1
    
    def _check_file_content(self, file_path: Path):
        """检查文件内容规范"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                lines = content.split('\n')
        except Exception as e:
            self.errors.append(f"无法读取文件 {file_path.relative_to(self.root_dir)}: {e}")
            return
        
        # 检查文件顶部注释
        self._check_file_header(file_path, lines)
        
        # 检查导入顺序
        self._check_import_order(file_path, lines)
        
        # 检查导入分组之间的空行（PEP 8 标准）
        self._check_import_blank_lines(file_path, lines)
        
        # 检查空行规范
        self._check_blank_lines(file_path, lines)
        
        # 检查行长度（PEP 8 标准：建议不超过99字符）
        self._check_line_length(file_path, lines)
        
        # 检查缩进（PEP 8 标准：使用4个空格，不使用Tab）
        self._check_indentation(file_path, lines)
        
        # 检查尾随空格（代码整洁）
        self._check_trailing_whitespace(file_path, lines)
        
        # 检查文件末尾换行符（PEP 8 标准）
        self._check_file_end_newline(file_path, content)
        
        # 检查TODO/FIXME注释
        self._check_todo_comments(file_path, lines)
        
        # 使用AST解析检查代码结构
        try:
            tree = ast.parse(content, filename=str(file_path))
            self._check_ast(file_path, tree)
        except SyntaxError as e:
            self.errors.append(f"语法错误 {file_path.relative_to(self.root_dir)}: {e}")
        except Exception as e:
            self.warnings.append(f"无法解析AST {file_path.relative_to(self.root_dir)}: {e}")
    
    def _check_file_header(self, file_path: Path, lines: List[str]):
        """检查文件头部注释"""
        # 跳过特殊文件
        if file_path.name in {'__init__.py'}:
            return
        
        # 检查前10行是否有注释
        has_comment = False
        for i, line in enumerate(lines[:10]):
            if line.strip().startswith('#'):
                has_comment = True
                break
        
        if not has_comment:
            self.warnings.append(f"文件缺少顶部注释: {file_path.relative_to(self.root_dir)}")
            self.stats['header_warnings'] += 1
    
    def _check_import_order(self, file_path: Path, lines: List[str]):
        """检查导入顺序（标准库/第三方库/本地应用）"""
        import_sections = []
        current_section = None
        
        for i, line in enumerate(lines, 1):
            stripped = line.strip()
            
            # 跳过注释和空行
            if not stripped or stripped.startswith('#'):
                continue
            
            # 检测导入语句
            if stripped.startswith('import ') or stripped.startswith('from '):
                # 判断导入类型
                if self._is_standard_library_import(stripped):
                    section = 'standard'
                elif self._is_third_party_import(stripped):
                    section = 'third_party'
                else:
                    section = 'local'
                
                if current_section is None:
                    current_section = section
                    import_sections.append((section, i))
                elif current_section != section:
                    # 检查顺序是否正确
                    if section == 'standard' and current_section in {'third_party', 'local'}:
                        self.warnings.append(
                            f"导入顺序错误 {file_path.relative_to(self.root_dir)}:第{i}行 "
                            f"(标准库应在第三方库和本地应用之前)"
                        )
                        self.stats['import_warnings'] += 1
                    elif section == 'third_party' and current_section == 'local':
                        self.warnings.append(
                            f"导入顺序错误 {file_path.relative_to(self.root_dir)}:第{i}行 "
                            f"(第三方库应在本地应用之前)"
                        )
                        self.stats['import_warnings'] += 1
                    current_section = section
                    import_sections.append((section, i))
            else:
                # 非导入语句，重置
                if current_section is not None:
                    current_section = None
    
    def _is_standard_library_import(self, import_line: str) -> bool:
        """判断是否为标准库导入"""
        # 提取模块名
        match = re.match(r'^(?:from|import)\s+([a-zA-Z0-9_]+)', import_line)
        if match:
            module = match.group(1)
            return module in STANDARD_LIBRARIES
        return False
    
    def _is_third_party_import(self, import_line: str) -> bool:
        """判断是否为第三方库导入"""
        # 常见的第三方库
        third_party_modules = {
            'yaml', 'tqdm', 'paramiko', 'openpyxl', 'xlsxwriter', 
            'requests', 'lxml', 'playwright', 'urllib3'
        }
        match = re.match(r'^(?:from|import)\s+([a-zA-Z0-9_]+)', import_line)
        if match:
            module = match.group(1)
            return module in third_party_modules
        return False
    
    def _check_import_blank_lines(self, file_path: Path, lines: List[str]):
        """检查导入分组之间的空行（PEP 8 标准：导入组之间应有空行）"""
        import_sections = []  # [(section, start_line, end_line), ...]
        current_section = None
        section_start = None
        
        for i, line in enumerate(lines, 1):
            stripped = line.strip()
            
            # 跳过注释
            if stripped.startswith('#'):
                continue
            
            # 检测导入语句
            if stripped.startswith('import ') or stripped.startswith('from '):
                # 判断导入类型
                if self._is_standard_library_import(stripped):
                    section = 'standard'
                elif self._is_third_party_import(stripped):
                    section = 'third_party'
                else:
                    section = 'local'
                
                if current_section is None:
                    current_section = section
                    section_start = i
                elif current_section != section:
                    # 分组切换，记录上一个分组
                    if section_start is not None:
                        import_sections.append((current_section, section_start, i - 1))
                    current_section = section
                    section_start = i
            else:
                # 非导入语句，结束当前分组
                if current_section is not None and section_start is not None:
                    import_sections.append((current_section, section_start, i - 1))
                    current_section = None
                    section_start = None
        
        # 处理最后一个分组
        if current_section is not None and section_start is not None:
            import_sections.append((current_section, section_start, len(lines)))
        
        # 检查分组之间是否有空行
        for idx in range(len(import_sections) - 1):
            current_end = import_sections[idx][2]
            next_start = import_sections[idx + 1][1]
            
            # 检查两个分组之间是否有空行
            if next_start - current_end == 1:
                # 没有空行，检查是否需要空行（不同分组之间需要空行）
                current_section = import_sections[idx][0]
                next_section = import_sections[idx + 1][0]
                if current_section != next_section:
                    self.warnings.append(
                        f"导入分组之间缺少空行 {file_path.relative_to(self.root_dir)}:第{current_end}行 "
                        f"(PEP 8 标准：不同导入组之间应有空行)"
                    )
                    self.stats['import_blank_line_warnings'] += 1
    
    def _check_trailing_whitespace(self, file_path: Path, lines: List[str]):
        """检查尾随空格（代码整洁）"""
        for i, line in enumerate(lines, 1):
            # 检查行尾是否有空格或Tab（排除空行，空行可能有意为空）
            if line.rstrip('\n\r') != line.rstrip('\n\r '):
                # 计算尾随空格数量
                trailing = len(line.rstrip('\n\r')) - len(line.rstrip('\n\r '))
                if trailing > 0:
                    self.warnings.append(
                        f"尾随空格 {file_path.relative_to(self.root_dir)}:第{i}行 "
                        f"(行尾有{trailing}个空格，建议删除以保持代码整洁)"
                    )
                    self.stats['trailing_whitespace_warnings'] += 1
    
    def _check_file_end_newline(self, file_path: Path, content: str):
        """检查文件末尾换行符（PEP 8 标准：文件末尾应有换行符）"""
        if content and not content.endswith('\n'):
            self.warnings.append(
                f"文件末尾缺少换行符 {file_path.relative_to(self.root_dir)} "
                f"(PEP 8 标准：文件末尾应有换行符)"
            )
            self.stats['file_end_newline_warnings'] += 1
    
    def _check_blank_lines(self, file_path: Path, lines: List[str]):
        """检查空行规范（最多连续2个空行，PEP 8 标准）"""
        consecutive_blank = 0
        for i, line in enumerate(lines, 1):
            if not line.strip():
                consecutive_blank += 1
                if consecutive_blank > 2:
                    self.warnings.append(
                        f"空行过多 {file_path.relative_to(self.root_dir)}:第{i}行 "
                        f"(连续{consecutive_blank}个空行，应最多2个，PEP 8 标准)"
                    )
                    self.stats['blank_line_warnings'] += 1
            else:
                consecutive_blank = 0
    
    def _check_line_length(self, file_path: Path, lines: List[str]):
        """检查行长度（PEP 8 标准：建议不超过99字符，允许最多120字符）"""
        MAX_LINE_LENGTH = 120  # 最大行长度
        WARN_LINE_LENGTH = 99   # 警告行长度（PEP 8 推荐）
        
        for i, line in enumerate(lines, 1):
            # 跳过注释行和空行
            stripped = line.strip()
            if not stripped or stripped.startswith('#'):
                continue
            
            # 检查行长度
            line_length = len(line.rstrip('\n\r'))
            if line_length > MAX_LINE_LENGTH:
                self.errors.append(
                    f"行长度过长 {file_path.relative_to(self.root_dir)}:第{i}行 "
                    f"({line_length}字符，超过最大限制{MAX_LINE_LENGTH}字符，PEP 8 标准)"
                )
                self.stats['line_length_errors'] += 1
            elif line_length > WARN_LINE_LENGTH:
                self.warnings.append(
                    f"行长度较长 {file_path.relative_to(self.root_dir)}:第{i}行 "
                    f"({line_length}字符，建议不超过{WARN_LINE_LENGTH}字符，PEP 8 标准)"
                )
                self.stats['line_length_warnings'] += 1
    
    def _check_indentation(self, file_path: Path, lines: List[str]):
        """检查缩进（PEP 8 标准：使用4个空格，不使用Tab）"""
        for i, line in enumerate(lines, 1):
            # 跳过空行
            if not line.strip():
                continue
            
            # 检查是否包含Tab字符
            if '\t' in line:
                self.errors.append(
                    f"使用Tab缩进 {file_path.relative_to(self.root_dir)}:第{i}行 "
                    f"(应使用4个空格，PEP 8 标准)"
                )
                self.stats['indentation_errors'] += 1
            
            # 检查缩进是否为4的倍数（对于有缩进的行）
            if line.startswith(' '):
                leading_spaces = len(line) - len(line.lstrip(' '))
                if leading_spaces % 4 != 0:
                    self.warnings.append(
                        f"缩进不规范 {file_path.relative_to(self.root_dir)}:第{i}行 "
                        f"({leading_spaces}个空格，应为4的倍数，PEP 8 标准)"
                    )
                    self.stats['indentation_warnings'] += 1
    
    def _check_todo_comments(self, file_path: Path, lines: List[str]):
        """检查TODO/FIXME/XXX/HACK注释（提醒开发者处理待办事项）"""
        todo_keywords = ['TODO', 'FIXME', 'XXX', 'HACK', 'NOTE', 'BUG']
        
        for i, line in enumerate(lines, 1):
            stripped = line.strip()
            if not stripped.startswith('#'):
                continue
            
            # 检查是否包含TODO等关键词
            for keyword in todo_keywords:
                if keyword in stripped.upper():
                    self.warnings.append(
                        f"待办注释 {file_path.relative_to(self.root_dir)}:第{i}行 "
                        f"(包含 {keyword}，请及时处理)"
                    )
                    self.stats['todo_warnings'] += 1
                    break
    
    def _check_ast(self, file_path: Path, tree: ast.AST):
        """使用AST检查代码结构"""
        visitor = CodeStyleASTVisitor(file_path, self)
        visitor.visit(tree)
        
        # 检查文档字符串
        self._check_docstrings(file_path, tree)
        
        # 检查异常处理
        self._check_exceptions(file_path, tree)
    
    def add_error(self, file_path: Path, line: int, message: str):
        """添加错误"""
        self.errors.append(f"{file_path.relative_to(self.root_dir)}:第{line}行 - {message}")
        self.stats['errors'] += 1
    
    def add_warning(self, file_path: Path, line: int, message: str):
        """添加警告"""
        self.warnings.append(f"{file_path.relative_to(self.root_dir)}:第{line}行 - {message}")
        self.stats['warnings'] += 1
    
    def _check_docstrings(self, file_path: Path, tree: ast.AST):
        """检查文档字符串（PEP 8 标准：公共函数和类应有docstring）"""
        visitor = DocstringChecker(file_path, self)
        visitor.visit(tree)
    
    def _check_exceptions(self, file_path: Path, tree: ast.AST):
        """检查异常处理（PEP 8 标准：避免bare except）"""
        visitor = ExceptionChecker(file_path, self)
        visitor.visit(tree)


class CodeStyleASTVisitor(ast.NodeVisitor):
    """AST访问器，检查代码规范（遵循 PEP 8 和 Pyright 标准）"""
    
    def __init__(self, file_path: Path, checker: CodeStyleChecker):
        self.file_path = file_path
        self.checker = checker
        self._context_stack = []  # 跟踪当前上下文（模块/类/函数）
    
    def visit_ClassDef(self, node: ast.ClassDef):
        """检查类名（驼峰命名，PEP 8 标准）"""
        class_name = node.name
        if not re.match(r'^[A-Z][a-zA-Z0-9]*$', class_name):
            self.checker.add_warning(
                self.file_path, node.lineno,
                f"类名命名不规范: {class_name} (应为驼峰命名 PascalCase，PEP 8 标准)"
            )
            self.checker.stats['class_warnings'] += 1
        
        # 检查类前是否有注释
        self._check_comment_before(node)
        
        # 进入类上下文
        self._context_stack.append('class')
        self.generic_visit(node)
        self._context_stack.pop()
    
    def visit_FunctionDef(self, node: ast.FunctionDef):
        """检查函数名（小写下划线，PEP 8 标准）"""
        func_name = node.name
        
        # 跳过Python特殊方法（如 __init__, __str__ 等）
        if func_name.startswith('__') and func_name.endswith('__'):
            self.generic_visit(node)
            return
        
        # 私有函数允许下划线开头（PEP 8 标准：单下划线表示内部使用）
        if func_name.startswith('_'):
            if not re.match(r'^_[a-z][a-z0-9_]*$', func_name):
                self.checker.add_warning(
                    self.file_path, node.lineno,
                    f"私有函数命名不规范: {func_name} (应为单下划线 + 小写下划线，PEP 8 标准)"
                )
                self.checker.stats['function_warnings'] += 1
        else:
            if not re.match(r'^[a-z][a-z0-9_]*$', func_name):
                self.checker.add_warning(
                    self.file_path, node.lineno,
                    f"函数名命名不规范: {func_name} (应为小写下划线 snake_case，PEP 8 标准)"
                )
                self.checker.stats['function_warnings'] += 1
        
        # 检查函数前是否有注释
        self._check_comment_before(node)
        
        # 进入函数上下文
        self._context_stack.append('function')
        self.generic_visit(node)
        self._context_stack.pop()
    
    def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef):
        """检查异步函数名（小写下划线，PEP 8 标准）"""
        func_name = node.name
        
        # 跳过Python特殊方法
        if func_name.startswith('__') and func_name.endswith('__'):
            self.generic_visit(node)
            return
        
        # 私有函数允许下划线开头（PEP 8 标准）
        if func_name.startswith('_'):
            if not re.match(r'^_[a-z][a-z0-9_]*$', func_name):
                self.checker.add_warning(
                    self.file_path, node.lineno,
                    f"私有异步函数命名不规范: {func_name} (应为单下划线 + 小写下划线，PEP 8 标准)"
                )
                self.checker.stats['function_warnings'] += 1
        else:
            if not re.match(r'^[a-z][a-z0-9_]*$', func_name):
                self.checker.add_warning(
                    self.file_path, node.lineno,
                    f"异步函数名命名不规范: {func_name} (应为小写下划线 snake_case，PEP 8 标准)"
                )
                self.checker.stats['function_warnings'] += 1
        
        # 检查函数前是否有注释
        self._check_comment_before(node)
        
        # 进入函数上下文
        self._context_stack.append('function')
        self.generic_visit(node)
        self._context_stack.pop()
    
    def visit_Assign(self, node: ast.Assign):
        """检查变量名（遵循 PEP 8：普通变量小写下划线，常量全大写）"""
        # 判断当前上下文：模块级、类内、函数内
        is_module_level = len(self._context_stack) == 0
        is_in_class = 'class' in self._context_stack
        is_in_function = 'function' in self._context_stack
        
        for target in node.targets:
            if isinstance(target, ast.Name):
                var_name = target.id
                
                # 跳过私有变量（单下划线开头，PEP 8 标准）
                if var_name.startswith('_'):
                    continue
                
                # 模块级变量检查（PEP 8 标准）
                if is_module_level:
                    # PEP 8：模块级变量可以是常量（全大写）或普通变量（小写下划线）
                    if re.match(r'^[A-Z][A-Z0-9_]*$', var_name):
                        # 全大写，符合常量规范（PEP 8）
                        pass
                    elif not re.match(r'^[a-z][a-z0-9_]*$', var_name):
                        # 既不是全大写也不是小写下划线，警告
                        self.checker.add_warning(
                            self.file_path, node.lineno,
                            f"模块级变量命名不规范: {var_name} (PEP 8 标准：应为小写下划线或全大写常量)"
                        )
                        self.checker.stats['variable_warnings'] += 1
                
                # 类属性检查（PEP 8 标准）
                elif is_in_class and not is_in_function:
                    # PEP 8：类属性通常使用小写下划线，常量可以使用全大写
                    if re.match(r'^[A-Z][A-Z0-9_]*$', var_name):
                        # 全大写，可能是类常量
                        pass
                    elif not re.match(r'^[a-z][a-z0-9_]*$', var_name):
                        # 既不是全大写也不是小写下划线，警告
                        self.checker.add_warning(
                            self.file_path, node.lineno,
                            f"类属性命名不规范: {var_name} (PEP 8 标准：应为小写下划线或全大写常量)"
                        )
                        self.checker.stats['variable_warnings'] += 1
                
                # 函数内部变量检查（PEP 8 标准：小写下划线）
                elif is_in_function:
                    # PEP 8：函数内部局部变量应使用小写下划线
                    if not re.match(r'^[a-z][a-z0-9_]*$', var_name):
                        # 检查是否为常量（全大写）
                        if not re.match(r'^[A-Z][A-Z0-9_]*$', var_name):
                            self.checker.add_warning(
                                self.file_path, node.lineno,
                                f"函数内部变量命名不规范: {var_name} (PEP 8 标准：应为小写下划线)"
                            )
                            self.checker.stats['variable_warnings'] += 1
        
        self.generic_visit(node)
    
    def _check_comment_before(self, node: ast.AST):
        """检查节点前是否有注释"""
        # 这个功能需要访问源代码，暂时跳过
        pass


class DocstringChecker(ast.NodeVisitor):
    """文档字符串检查器（PEP 8 标准）"""
    
    def __init__(self, file_path: Path, checker: CodeStyleChecker):
        self.file_path = file_path
        self.checker = checker
    
    def visit_ClassDef(self, node: ast.ClassDef):
        """检查类是否有文档字符串"""
        if not ast.get_docstring(node):
            self.checker.add_warning(
                self.file_path, node.lineno,
                f"类缺少文档字符串: {node.name} (PEP 8 标准：公共类应有docstring)"
            )
            self.checker.stats['docstring_warnings'] += 1
        self.generic_visit(node)
    
    def visit_FunctionDef(self, node: ast.FunctionDef):
        """检查函数是否有文档字符串"""
        # 跳过私有函数（单下划线开头）和特殊方法
        if node.name.startswith('__') and node.name.endswith('__'):
            self.generic_visit(node)
            return
        
        # 跳过私有函数（单下划线开头）
        if node.name.startswith('_'):
            self.generic_visit(node)
            return
        
        # 检查公共函数是否有文档字符串
        if not ast.get_docstring(node):
            self.checker.add_warning(
                self.file_path, node.lineno,
                f"公共函数缺少文档字符串: {node.name} (PEP 8 标准：公共函数应有docstring)"
            )
            self.checker.stats['docstring_warnings'] += 1
        
        self.generic_visit(node)
    
    def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef):
        """检查异步函数是否有文档字符串"""
        # 异步函数与普通函数使用相同的检查逻辑
        if node.name.startswith('__') and node.name.endswith('__'):
            self.generic_visit(node)
            return
        
        if node.name.startswith('_'):
            self.generic_visit(node)
            return
        
        if not ast.get_docstring(node):
            self.checker.add_warning(
                self.file_path, node.lineno,
                f"公共异步函数缺少文档字符串: {node.name} (PEP 8 标准：公共函数应有docstring)"
            )
            self.checker.stats['docstring_warnings'] += 1
        
        self.generic_visit(node)


class ExceptionChecker(ast.NodeVisitor):
    """异常处理检查器（PEP 8 标准）"""
    
    def __init__(self, file_path: Path, checker: CodeStyleChecker):
        self.file_path = file_path
        self.checker = checker
    
    def visit_ExceptHandler(self, node: ast.ExceptHandler):
        """检查异常处理（PEP 8 标准：避免bare except）"""
        # 检查是否为 bare except（没有指定异常类型）
        if node.type is None:
            self.checker.add_warning(
                self.file_path, node.lineno,
                "使用 bare except (PEP 8 标准：应指定具体异常类型，如 except Exception:)"
            )
            self.checker.stats['exception_warnings'] += 1
        
        self.generic_visit(node)


def main():
    """主函数"""
    import sys
    
    root_dir = sys.argv[1] if len(sys.argv) > 1 else "."
    
    print("=" * 80)
    print("代码规范检查工具")
    print("=" * 80)
    print()
    
    checker = CodeStyleChecker(root_dir)
    errors, warnings, stats = checker.check_all()
    
    # 输出结果
    if errors:
        print("❌ 错误:")
        for error in errors:
            print(f"  {error}")
        print()
    
    if warnings:
        print("⚠️  警告:")
        for warning in warnings:
            print(f"  {warning}")
        print()
    
    # 输出统计信息
    print("📊 统计信息:")
    for key, value in sorted(stats.items()):
        print(f"  {key}: {value}")
    print()
    
    # 总结
    total_issues = len(errors) + len(warnings)
    if total_issues == 0:
        print("✅ 代码规范检查通过！")
        return 0
    else:
        print(f"❌ 发现 {len(errors)} 个错误，{len(warnings)} 个警告")
        return 1


if __name__ == "__main__":
    exit(main())

