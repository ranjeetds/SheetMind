"""
Setup script for SheetMind.

AI-powered Excel automation tool with natural language processing.
"""

from pathlib import Path
from setuptools import find_packages, setup

# Read README for long description
readme_path = Path(__file__).parent / "README.md"
long_description = readme_path.read_text(encoding="utf-8") if readme_path.exists() else ""

# Read requirements
requirements_path = Path(__file__).parent / "requirements.txt"
requirements = []
if requirements_path.exists():
    with open(requirements_path, "r", encoding="utf-8") as f:
        requirements = [line.strip() for line in f if line.strip() and not line.startswith("#")]

setup(
    name="sheetmind",
    version="0.1.0",
    author="SheetMind Contributors",
    author_email="hello@sheetmind.dev",
    description="AI-powered Excel automation tool with natural language processing",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/sheetmind",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: End Users/Desktop",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Scientific/Engineering :: Artificial Intelligence",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=7.4.0",
            "pytest-asyncio>=0.21.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "mypy>=1.0.0",
        ],
        "full": [
            "matplotlib>=3.7.0",
            "seaborn>=0.12.0",
            "plotly>=5.15.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "sheetmind=src.main:cli",
        ],
    },
    include_package_data=True,
    package_data={
        "ui.web": ["templates/*.html", "static/*"],
    },
    zip_safe=False,
    keywords="excel ai automation nlp spreadsheet",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/sheetmind/issues",
        "Source": "https://github.com/yourusername/sheetmind",
        "Documentation": "https://sheetmind.readthedocs.io",
    },
) 