from setuptools import setup
from md2paper.version import __version__
# repeated requirements is error-prone
REQUIREMENTS = [
    'python-docx==0.8.11',
    'latex2mathml==3.63.3',
    'Markdown==3.3.6',
    'beautifulsoup4>=4.9.3',
    'bibtexparser==1.2.0',
    'Pillow==9.0.0'
]

# python-docx和bibtexparser在wasm环境下无法直接安装(缺少直接的wheel)，
# 在安装md2paper前应该手动安装以上依赖项
# 选项之一是使用我们手动打包的.whl

setup(
    name = 'md2paper',
    version = __version__,
    url = 'https://github.com/tzy15368/md2paper',
    author = 'indigo15, KZNS',
    author_email= 'tzy15368@outlook.com',
    packages = ['md2paper'],
    install_requires = REQUIREMENTS
)