from setuptools import setup, find_packages

setup(
    name='slide2pdf',
    version='0.1',
    packages=find_packages(),
    install_requires=[
        'comtypes',
    ],
    entry_points={
        'console_scripts': [
            'slide2pdf=slide2pdf.cli:main',
        ],
    },
    author='Arunmozhi Varman',
    author_email='amv.k.2712005@gmail.com',
    description='CLI tool to convert all PowerPoint files in a folder to PDF using MS PowerPoint',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/yourusername/slide2pdf', 
    classifiers=[
        'Programming Language :: Python :: 3',
        'Operating System :: Microsoft :: Windows',
        'Intended Audience :: Developers',
        'Topic :: Utilities',
    ],
    python_requires='>=3.6',
)
