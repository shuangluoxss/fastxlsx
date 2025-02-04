# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

# import fastxlsx
# import os
# import sys
# import shutil
# from pathlib import Path

# # 源存根目录（存放.pyi文件）
# stub_source = Path("../")
# # 临时存根目录（存放.py文件）
# stub_dest = Path("_temp_stubs")

# # 清理并创建临时目录
# if stub_dest.exists():
#     shutil.rmtree(stub_dest)
# stub_dest.mkdir()

# # 复制并重命名.pyi文件到临时目录
# for pyi_path in stub_source.glob("*.pyi"):
#     relative_path = pyi_path.relative_to(stub_source)
#     py_path = stub_dest / relative_path.with_suffix(".py")
#     py_path.parent.mkdir(parents=True, exist_ok=True)
#     shutil.copy(pyi_path, py_path)

# # 将临时目录添加到模块搜索路径的最前面
# sys.path.insert(0, str(stub_dest.resolve()))

project = "FastXLSX"
copyright = "2025, shuangluoxss"
author = "shuangluoxss"
# version = fastxlsx.version()

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = [
    "sphinx.ext.autodoc",  # 自动生成 API 文档
    "sphinx.ext.viewcode",  # 添加源代码链接
    "sphinx.ext.napoleon",  # 支持 Google 和 NumPy 风格的文档字符串
    # "sphinx_autodoc_typehints",  # 支持类型提示
    "myst_parser",  # 支持 Markdown 格式
]
autodoc_typehints = "both"  # 显示类型注解
autodoc_stub_files = False  # 启用 .pyi 文件解析


templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = "sphinx_rtd_theme"
html_theme_options = {
    "version_selector": True,  # 在侧边栏显示版本号
    # "version": version,  # 设置版本号
}
html_context = {
    "display_github": True,
    "github_user": "shuangluoxss",
    "github_repo": "fastxlsx",
    "github_version": "main",
    "conf_py_path": "/doc/",
}
html_static_path = ["_static"]
