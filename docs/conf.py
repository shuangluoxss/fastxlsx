# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

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
    # "myst_parser",  # 支持 Markdown 格式
]
autodoc_typehints = "both"  # 显示类型注解
autodoc_stub_files = False  # 启用 .pyi 文件解析


templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = "sphinx_rtd_theme"
html_context = {
    "display_github": True,
    "github_user": "shuangluoxss",
    "github_repo": "fastxlsx",
    "github_version": "main",
    "conf_py_path": "/docs/",
}
html_static_path = ["_static"]
