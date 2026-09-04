from setuptools import setup, find_packages

setup(
    name="catia-ai-bridge",
    version="1.0.0",
    author="DRISSI AMJAD",
    description="CATIA V5 R21 AI Studio & Automation Bridge",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/AM-DR/Catia-AI-bridge",
    py_modules=["app", "run_app"],
    install_requires=[
        "streamlit>=1.30.0",
        "pycatia>=0.5.8",
        "pywin32>=306",
        "langchain>=0.2.0",
        "langchain-core>=0.2.0",
        "langchain-community>=0.2.0",
        "langchain-openai>=0.1.0",
        "langchain-anthropic>=0.1.0",
        "langchain-ollama>=0.1.0",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: Microsoft :: Windows",
        "License :: OSI Approved :: MIT License",
    ],
    python_requires=">=3.10",
)
