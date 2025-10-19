from setuptools import setup

setup(
    name="similarity-answer-matcher",
    version="0.1.0",
    install_requires=[
        "streamlit>=1.24.0,<1.25.0",
        "pandas>=2.0.3,<2.1.0",
        "numpy>=1.24.3,<1.25.0",
        "sentence-transformers>=2.2.2,<2.3.0",
        "transformers>=4.30.2,<4.31.0",
        "torch>=2.0.0,<2.1.0",
        "torchvision>=0.15.0,<0.16.0"
    ],
    python_requires=">=3.9,<3.14",
)